/** 1日の長さ */
const DAY_MS = 24 * 60 * 60 * 1000

/** タイムゾーンの補正 */
const TZ_OFFSET_MS = 9 * 60 * 60 * 1000

/** 一度に保持する最大行数 */
const MAX_DAYS_LENGTH = 31

/**
 * 与えられたDateを日本時間での日単位に切り下げる
 * @param d 入力値 - 省略時は現在時刻
 * @returns 新しいDateオブジェクト
 */
function floorTime(d?: Date): number {
  return (
    Math.floor((d ? d.getTime() : Date.now() - TZ_OFFSET_MS) / DAY_MS) *
      DAY_MS +
    TZ_OFFSET_MS
  )
}

export function updateDateRows(): void {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet()
  const sheets = spreadsheet.getSheets()
  const today = new Date(floorTime())

  for (const sheet of sheets) {
    const lastRow = sheet.getLastRow()
    const lastColumn = sheet.getLastColumn()

    // データが入ってないシートはスキップ
    if (lastRow === 0) {
      continue
    }

    // --- 挿入パート ---

    const lastValue = sheet.getRange(lastRow, 1).getValue()

    // 最後の行が日付値でなければスキップ
    if (!(lastValue instanceof Date)) {
      continue
    }

    // 追加するべき行数の決定（MAX_DAYS_LENGTH日分先の日付〜最後の行に存在する日付の差）
    const lengthToInsert = Math.ceil(
      (today.getTime() + MAX_DAYS_LENGTH * DAY_MS - lastValue.getTime()) /
        DAY_MS
    )

    // 1行以上追加がある
    if (lengthToInsert > 0) {
      const rows: Date[][] = []

      // 追加する日付データの生成
      for (let i = 0; i < lengthToInsert; i++) {
        rows.push([new Date(lastValue.getTime() + DAY_MS * (i + 1))])
      }

      // データを挿入する
      sheet.insertRowsAfter(lastRow, rows.length)

      sheet.getRange(lastRow + 1, 1, rows.length, 1).setValues(rows)

      // 書式と数式をコピーする
      if (lastColumn > 1) {
        const sourceRange = sheet.getRange(lastRow, 2, 1, lastColumn - 1)

        // 参照を含む数式がある列のみ数式をコピーする
        const [formulars] = sourceRange.getFormulas()
        for (let i = 0; i < formulars.length; i++) {
          if (formulars[i].match(/^=.*[A-Z][0-9]/)) {
            sheet
              .getRange(lastRow, i + 2)
              .copyTo(
                sheet.getRange(lastRow + 1, i + 2, rows.length, 1),
                SpreadsheetApp.CopyPasteType.PASTE_FORMULA,
                false
              )
          }
        }

        // 書式は全てコピーする
        const targetRange = sheet.getRange(
          lastRow + 1,
          2,
          rows.length,
          lastColumn - 1
        )
        sourceRange.copyTo(
          targetRange,
          SpreadsheetApp.CopyPasteType.PASTE_FORMAT,
          false
        )
        sourceRange.copyTo(
          targetRange,
          SpreadsheetApp.CopyPasteType.PASTE_CONDITIONAL_FORMATTING,
          false
        )
      }
    }

    // --- 削除パート ---
    // もとあった部分の値の取得
    const values = sheet.getRange(1, 1, lastRow, 1).getValues()

    /** 削除された行ぶんの補正 */
    let offset = 0

    for (let i = 0; i < values.length; i++) {
      const v = values[i][0]

      // 日付型の値が入っているか
      if (v instanceof Date && v < today) {
        // もう過ぎた日付であるなら削除
        sheet.deleteRow(i + offset + 1)
        offset--
      }
    }
  }
}
