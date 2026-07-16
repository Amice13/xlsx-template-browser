import type JSZip from 'jszip'
import type { Relations } from '../types/relations'
import { parser } from './global-helpers'
import type { Drawings } from '../types/drawings'

const serializeXml = (content: string): string => (`
<xml xmlns:v="urn:schemas-microsoft-com:vml"
     xmlns:o="urn:schemas-microsoft-com:office:office"
     xmlns:x="urn:schemas-microsoft-com:office:excel">

  <o:shapelayout v:ext="edit">
    <o:idmap v:ext="edit" data="1"/>
  </o:shapelayout>

  <v:shapetype id="_x0000_t202" coordsize="21600,21600" o:spt="202" path="m,l,21600r21600,l21600,xe">
    <v:stroke joinstyle="miter"/>
    <v:path gradientshapeok="t" o:connecttype="rect"/>
  </v:shapetype>
  ${content}
</xml>
`)

const createShape = ({
  id,
  row,
  column
}: {
  id: number
  row: number
  column: number
}): string => {
  return `<v:shape id="_x0000_s${String(1025 + id)}" type="#_x0000_t202"
  style="position:absolute;width:10pt;height:10pt;visibility:hidden"
  fillcolor="none" strokecolor="none">
  <v:fill color2="infoBackground [80]"/>
  <v:shadow color="none [81]" obscured="t"/>
  <v:path o:connecttype="none"/>
  <v:textbox style="mso-direction-alt:auto">
  <div style="text-align:left"/>
  </v:textbox>
  <x:ClientData ObjectType="Note">
  <x:MoveWithCells/>
  <x:SizeWithCells/>
  <x:Anchor>${column},0,${row},0,${column + 1},0,${row + 2},0</x:Anchor>
  <x:AutoFill>False</x:AutoFill>
  <x:Row>${row}</x:Row>
  <x:Column>${column}</x:Column>
  </x:ClientData>
  </v:shape>`
}

export const getDrawings = async ({
  id,
  relations,
  xlsx
}: {
  id: string
  relations: Relations
  xlsx: JSZip
}): Promise<Drawings> => {
  const drawingRelation = relations.getElements('vmlDrawing')
  const guessedFilename = `../drawings/vmlDrawing${id.replace('rId', '')}.vml`
  const noRelation = drawingRelation.length === 0
  const filename = (drawingRelation[0]?.target ?? guessedFilename).replace('..', 'xl')
  const unique = new Set<string>()
  let isDirty = false
  const drawings: Array<{ row: string, column: string }> = []
  const xmlText = await xlsx.file(filename)?.async('string')
  if (xmlText !== undefined) {
    const vmlDoc = parser.parseFromString(xmlText, 'application/xml')
    const shapes = vmlDoc.getElementsByTagName('v:shape')
    for (const shape of shapes) {
      const row = shape.getElementsByTagName('x:Row')[0]?.textContent
      const column = shape.getElementsByTagName('x:Column')[0]?.textContent
      if (row === undefined || column === undefined) continue
      const rowValue = String(parseInt(row) + 1)
      const columnValue = String(parseInt(column) + 1)
      drawings.push({ row: rowValue, column: columnValue })
    }
  }

  const add = ({ row, column }: { row: string, column: string }): void => {
    const ref = `${row}:${column}`
    if (unique.has(ref)) return
    drawings.push({ row, column })
    unique.add(ref)
    isDirty = true
  }

  const save = (): void => {
    if (!isDirty) return
    if (noRelation) {
      relations.add({
        target: filename.replace('xl', '..'),
        type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing'
      })
    }
    const content = drawings.map((drawing, i) => {
      return createShape({
        id: i,
        row: parseInt(drawing.row) - 1,
        column: parseInt(drawing.column) - 1
      })
    }).join('')
    const data = serializeXml(content)
    xlsx.file(filename, data)
  }
  return {
    add,
    save
  }
}
