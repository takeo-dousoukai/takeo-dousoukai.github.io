/* ============================================================
 * 最終出欠確認用の Googleフォーム
 *
 * 使い方（1回だけ）
 *   1. スプレッドシートの「拡張機能 → Apps Script」を開く
 *   2. 左の「ファイル ＋」→「スクリプト」で FinalForm という名前のファイルを作り、これを貼る
 *   3. 上の関数選択で createFinalForm を選んで「実行」（初回は権限の許可が出ます）
 *   4. 実行ログに出る「回答用URL」をサイトと案内文に使う
 *
 * フォームの回答は
 *   ・スプレッドシートの「フォームの回答」シートにそのまま残る（原本）
 *   ・同時に 名簿 / 受付ログ にも反映される → サイトの人数・幹事ページもそのまま動く
 * ============================================================ */

const FINAL_FORM_TITLE = "武雄高校 同窓会 最終出欠確認（12/28）";
const FINAL_Q = {
  cls: "卒業時のクラス",
  name: "お名前（卒業時の姓名）",
  oldName: "今の姓（苗字が変わった方のみ）",
  attendance: "12月28日（月）の同窓会に",
  line: "LINEの名前",
  email: "メールアドレス",
  tel: "電話番号",
  message: "ひとこと",
};

function createFinalForm() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const props = PropertiesService.getScriptProperties();
  if (props.getProperty("FINAL_FORM_ID")) {
    const f = FormApp.openById(props.getProperty("FINAL_FORM_ID"));
    Logger.log("作成済みです。回答用URL: " + f.getPublishedUrl());
    return;
  }

  const form = FormApp.create(FINAL_FORM_TITLE)
    .setDescription(
      "佐賀県立武雄高等学校 平成27年3月卒業（3年1組〜7組）の同窓会です。\n\n" +
      "日時：12月28日（月）夜\n会場：武雄センチュリーホテル\n先生方もお招きします。\n\n" +
      "この回答で人数を確定して会場に伝えます。前のアンケートに答えた方も、もう一度ご回答ください。\n" +
      "来られない方も「参加しない」でご回答いただけると助かります。\n" +
      "締め切り：11月15日（日）")
    .setCollectEmail(false)
    .setAllowResponseEdits(true)
    .setLimitOneResponsePerUser(false)
    .setConfirmationMessage(
      "回答ありがとうございました。\n会費が決まったら、PayPayでの事前入金のご案内をお送りします。" +
      "入金を確認できた方から、参加者用のLINEグループに招待します。");

  form.addMultipleChoiceItem().setTitle(FINAL_Q.cls).setRequired(true)
    .setChoiceValues(["1組", "2組", "3組", "4組", "5組", "6組", "7組"]);
  form.addTextItem().setTitle(FINAL_Q.name).setRequired(true)
    .setHelpText("名簿と照らし合わせるので、卒業時のお名前で入力してください（例：武雄 太郎）");
  form.addTextItem().setTitle(FINAL_Q.oldName)
    .setHelpText("結婚などで苗字が変わった方は、今の苗字を入れてください");
  form.addMultipleChoiceItem().setTitle(FINAL_Q.attendance).setRequired(true)
    .setChoiceValues(["参加する", "参加しない"]);
  form.addSectionHeaderItem().setTitle("連絡先")
    .setHelpText("どれか1つは入れてください。幹事だけが見て、同窓会の連絡にだけ使います。会が終わったら削除します。");
  form.addTextItem().setTitle(FINAL_Q.line);
  form.addTextItem().setTitle(FINAL_Q.email);
  form.addTextItem().setTitle(FINAL_Q.tel);
  form.addParagraphTextItem().setTitle(FINAL_Q.message)
    .setHelpText("近況や先生への伝言、会いたい人など（任意）");

  form.setDestination(FormApp.DestinationType.SPREADSHEET, ss.getId());
  ScriptApp.newTrigger("onFinalFormSubmit").forSpreadsheet(ss).onFormSubmit().create();
  props.setProperty("FINAL_FORM_ID", form.getId());

  Logger.log("回答用URL: " + form.getPublishedUrl());
  Logger.log("編集用URL: " + form.getEditUrl());
}

/** フォーム送信のたびに 名簿 / 受付ログ に反映する */
function onFinalFormSubmit(e) {
  const v = (title) => ((e.namedValues[title] || [""])[0] || "").toString().trim();
  const line = v(FINAL_Q.line), email = v(FINAL_Q.email), tel = v(FINAL_Q.tel);
  const contact = [line && "LINE: " + line, email && "メール: " + email, tel && "電話: " + tel]
    .filter(Boolean).join(" / ");
  submitRsvp_({
    cls: v(FINAL_Q.cls).replace("組", ""),
    name: v(FINAL_Q.name),
    oldName: v(FINAL_Q.oldName),
    attendance: v(FINAL_Q.attendance) === "参加する" ? "参加" : "不参加",
    contact: contact,
    dates: "最終確認（Googleフォーム）",
    message: v(FINAL_Q.message),
  });
}
