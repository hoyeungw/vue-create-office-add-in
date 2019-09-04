import { Mat } from 'veho'
import { MatX } from 'xbrief'

let distinct = async (context) => {
  let range = context.workbook.getSelectedRange()
  range.load('getAbsoluteResizedRange')
  range.load('getCell')
  // range.format.fill.color = 'green'
  const matrix = Mat.ini(3, 3, (i, j) => i + j + 1)
  MatX.xBrief(matrix).wL()
  await context.sync().then(() => {
    // range.getCell(0, 0).values = [['here we are']]
    range.getAbsoluteResizedRange(3, 3).values = matrix
  })
}

export {
  distinct,
}
