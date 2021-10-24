import { PropertiesKey } from '.'

export function triggerScheduleReminder(): void {
  const now = new Date()
  const slackWebhookUrl = PropertiesService.getScriptProperties().getProperty(
    PropertiesKey.slackWebhookUrl
  )

  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet()
  const sheet = spreadsheet.getSheets()[0]

  const lastColumn = sheet.getLastColumn()
  const lastRow = sheet.getLastRow()

  // ヘッダ行を見つける
  const [headers] = sheet.getRange(1, 1, 1, lastColumn).getValues()

  let data
  for (let i = 1; i < lastRow; i++) {
    // prettier-ignore
    [data] = sheet.getRange(i, 1, 1, lastColumn).getValues()

    // 日付値が1列目に入っている行を見つける
    if (data[0] instanceof Date) {
      // きょうの日付の行を探す（現在時刻の24時間前後以内 & 「日」の値が同じ）
      if (
        Math.abs(data[0].getTime() - now.getTime()) < 1000 * 60 * 60 * 24 &&
        data[0].getDate() === now.getDate()
      ) {
        break
      }
    }
  }

  let eventColumn = 0
  for (; eventColumn < headers.length; eventColumn++) {
    if (headers[eventColumn].match(/活動/)) {
      // 「活動」を含むヘッダ列を予定列として扱う
      break
    }
  }

  // 予定列が空の場合は何もしない
  if (!data[eventColumn] || data[eventColumn] === '') {
    console.info('本日の活動予定はありません。')
    return
  }

  /** 参加者の配列 */
  const attendees: {
    name: string
    mode: 'attend' | 'absent' | 'other'
    message?: string
  }[] = []

  // 参加者の出欠を振り分ける
  for (let i = eventColumn + 1; i < data.length; i++) {
    /** 参加者ごとのoverviewシートの値 */
    const message = data[i]

    // 未回答
    if (message.length === 0 || message === '-') {
      continue
    }
    // 欠席（☓・☓☓は現状同一とする）
    else if (message.match(/^[x✕✗☓]+$/)) {
      attendees.push({
        name: headers[i],
        mode: 'absent',
      })
    }
    // 出席
    else if (message.match(/^[◯○◎]$/)) {
      attendees.push({ name: headers[i], mode: 'attend' })
    }
    // △指定
    else {
      attendees.push({ name: headers[i], mode: 'other', message: message })
    }
  }

  // Slackに送信する
  UrlFetchApp.fetch(slackWebhookUrl, {
    method: 'post',
    payload: JSON.stringify({
      text: `活動予定のリマインド: ${now.getMonth() + 1}/${now.getDate()} ${
        data[eventColumn]
      }`,
      blocks: [
        {
          type: 'divider',
        },
        {
          type: 'section',
          text: {
            type: 'mrkdwn',
            text: `*${now.getMonth() + 1}/${now.getDate()} ${data[eventColumn]}*
            
            :o: ${attendees
              .filter(a => a.mode === 'attend')
              .map(a => a.name)
              .join('・')}
            ${attendees
              .filter(a => a.mode === 'other')
              .map(a => `:thinking_face: ${a.name} - ${a.message}`)
              .join('\n')}
            :x: ${attendees
              .filter(a => a.mode === 'absent')
              .map(a => a.name)
              .join('・')}`.replace(/^\s+(?=\S)/gm, ''), // 先頭の空白を削除
          },
          accessory: {
            type: 'image',
            image_url:
              'https://api.slack.com/img/blocks/bkb_template_images/notifications.png',
            alt_text: '（カレンダーのアイコン）',
          },
        },
        {
          type: 'context',
          elements: [
            {
              type: 'mrkdwn',
              text: '当日の連絡・予定変更はこのメッセージにスレッドで付け足してください。',
            },
            {
              type: 'mrkdwn',
              text: '<https://atnd.tkbneu.net|予定調査を開く>',
            },
          ],
        },
      ],
    }),
  })

  console.info('リマインダを投稿しました。')
}
