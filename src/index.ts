export enum PropertiesKey {
  slackWebhookUrl = 'slack_webhook_url',
  slackBotUserOAuthToken = 'slack_bot_user_oauth_token',
}

/**
 * WebUIを表示する
 */
export function doGet(): GoogleAppsScript.HTML.HtmlOutput {
  // ここでユーザーを取得しないと必要な権限が取得されないまま動いてしまう
  const user = Session.getActiveUser()
  console.log(user.getEmail())

  return HtmlService.createTemplateFromFile('config').evaluate()
}

/**
 * WebUIからの設定更新をキャッチする
 */
export function handleUpdateFromWebUI({
  slackWebhookUrl,
  slackBotUserOAuthToken,
}: {
  slackWebhookUrl: string
  slackBotUserOAuthToken: string
}): void {
  PropertiesService.getScriptProperties().setProperties({
    [PropertiesKey.slackWebhookUrl]: slackWebhookUrl,
    [PropertiesKey.slackBotUserOAuthToken]: slackBotUserOAuthToken,
  })
}
