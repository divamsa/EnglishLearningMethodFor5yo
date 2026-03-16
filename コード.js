// ==========================================
// 1. Gemini APIキーの設定
// ==========================================
const GEMINI_API_KEY = PropertiesService.getScriptProperties().getProperty("GEMINI_API_KEY"); // ← 必ずご自身のAPIキーに書き換えてください

// ==========================================
// 2. ビジネス日本語 100選
// ==========================================
const BUSINESS_JP_LIST = [
  "先日の会議の件でフォローアップさせてください。", "詳細についてもっと詳しく教えていただけますか？", "できるだけ早く折り返しご連絡いたします。", "来週、お電話にてお話しするお時間をいただけますか？", "迅速なご返信をいただき、ありがとうございます。",
  "あいにくその時間は先約がございます。", "会議の時間を午後3時に後ろ倒しすることは可能でしょうか？", "こちらから妥協案を提示させてください。", "それは非常に妥当な取引条件だと思います。", "競争力を維持するためにコストを削減する必要があります。",
  "ご一緒にお仕事ができるのを楽しみにしております。", "ご不明な点がございましたら、お気軽にお知らせください。", "本日の議論はこのあたりで締めくくりましょう。", "その点については私も同意見です。", "その点について、もう少し詳しく説明していただけますか？",
  "この締め切りについては交渉の余地がありません。", "時間が残り少なくなってまいりました。", "来週の月曜日にまた状況を確認し合いましょう。", "ただちにその件に対応いたします。", "今回の件で最も重要なポイントは何でしょうか？"
];

// ==========================================
// 3. 準備用関数（C列に例文をセット）
// ==========================================
function setupSheet() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  sheet.clear();
  sheet.setName("学習シート"); // 分かりやすいようにシート名を変更

  const headers = ["A", "B", "【C列】ビジネス日本語", "【D列】5歳児の日本語 (生徒入力)", "【E列】英訳 (生徒入力)", "【F列】AI模範解答", "【G列】AI添削フィードバック"];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange("C1:G1").setBackground("#4A86E8").setFontColor("white").setFontWeight("bold");
  
  sheet.setColumnWidth(3, 250); 
  sheet.setColumnWidth(4, 250); 
  sheet.setColumnWidth(5, 250); 
  sheet.setColumnWidth(6, 300); 
  sheet.setColumnWidth(7, 350); 

  sheet.getRange("F:G").setWrap(true).setVerticalAlignment("top");

  const outputData = BUSINESS_JP_LIST.map(jp => [jp]);
  sheet.getRange(2, 3, outputData.length, 1).setValues(outputData);
  
  SpreadsheetApp.getActiveSpreadsheet().toast("準備完了！D列とE列に入力してください。", "完了");
}

// ==========================================
// ★新規：使い方シート（マニュアル）を作成する関数
// ==========================================
function createManualSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let manualSheet = ss.getSheetByName("使い方");

  // すでに「使い方」シートがあれば内容をクリア、なければ先頭に新規作成
  if (!manualSheet) {
    manualSheet = ss.insertSheet("使い方", 0);
  } else {
    manualSheet.clear();
  }

  const manualText = [
    ["【生徒さん向け】ビジネス英語「シンプル変換」トレーニングの使い方"],
    [""],
    ["このシートは、難しいビジネス日本語を「一度5歳児でもわかるレベルに噛み砕いてから英語にする」という、英語脳を作るための特別なトレーニングツールです。"],
    ["直訳しようとすると言葉に詰まってしまうビジネス表現も、このステップを踏むことで驚くほど簡単に英語で言えるようになります。専属のAIコーチ（Gemini）が24時間あなたの解答を添削してくれます！"],
    [""],
    ["📝 基本の進め方（3ステップ）"],
    ["Step 1：【C列】の「ビジネス日本語」を読む"],
    ["まずはC列に書かれている実際のビジネスシーンで使う表現を確認します。（例：「こちらから妥協案を提示させてください。」）"],
    [""],
    ["Step 2：【D列】に「5歳児の日本語」を入力する"],
    ["C列の言葉を、5歳児でも理解できるくらい簡単な日本語に言い換えてD列に入力します。ここが一番重要なトレーニングです！\n（例：「お互いが『いいね』って言えるアイデアを出してもいい？」）"],
    [""],
    ["Step 3：【E列】に「英訳」を入力する（Enterを押す）"],
    ["D列で作った「簡単な日本語」をベースにして、知っている英単語でE列に英文を入力します。\n（例：「Can I give you an idea we both like?」）"],
    [""],
    ["💡 AIからのフィードバック（自動表示）"],
    ["E列に英語を入力してEnterを押すと、数秒後に以下の内容が自動で表示されます！"],
    ["・【F列】AI模範解答：\nあなたが作った「5歳児日本語」に対する自然な英語と、元の「ビジネス日本語」に対するプロフェッショナルな英語の2パターンを提示します。"],
    ["・【G列】AI添削フィードバック：\nあなたがE列に書いた英語に対して、AIのGemini先生が「ここが良かった！」「こうするともっと自然になるよ！」と日本語で優しく添削・アドバイスをしてくれます。"],
    [""],
    ["⚠️ ご利用時の注意点"],
    ["・D列（5歳児の日本語）は必ず先に入力してください。空欄のままE列（英語）を入力すると、AIから「先に入力してね！」と注意されてしまいます。"],
    ["・入力を間違えた時は？ E列の英語を消去（Delete）すると、F列とG列のAIの回答も綺麗にリセットされます。何度でも書き直して挑戦できます。"],
    ["・AIが考え中（「🔄 Geminiが添削中...」）の時は、結果が出るまで数秒だけそのままお待ちください。"]
  ];

  manualSheet.getRange(1, 1, manualText.length, 1).setValues(manualText);

  // 見た目を綺麗に整える
  manualSheet.setColumnWidth(1, 800); // A列を広くして見やすく
  manualSheet.getRange("A:A").setWrap(true).setVerticalAlignment("middle");

  // 見出しの色付けとフォント調整
  manualSheet.getRange("A1").setFontSize(14).setFontWeight("bold").setBackground("#4A86E8").setFontColor("white");
  manualSheet.getRange("A6").setFontSize(12).setFontWeight("bold").setBackground("#e6fffa");
  manualSheet.getRange("A15").setFontSize(12).setFontWeight("bold").setBackground("#e6fffa");
  manualSheet.getRange("A20").setFontSize(12).setFontWeight("bold").setBackground("#fff4e6");

  ss.setActiveSheet(manualSheet); // 作成したシートを開く
  SpreadsheetApp.getActiveSpreadsheet().toast("「使い方」シートを作成しました！", "完了");
}

// ==========================================
// 🌟自動実行トリガーを設定する関数
// ==========================================
function createTrigger() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet();
  const triggers = ScriptApp.getProjectTriggers();
  for (let i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'autoGeminiFeedback') {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
  ScriptApp.newTrigger('autoGeminiFeedback')
    .forSpreadsheet(sheet)
    .onEdit()
    .create();
  SpreadsheetApp.getUi().alert("✅ 自動実行の設定が完了しました！");
}

// ==========================================
// 4. E列に入力されたらGeminiを呼び出す関数
// ==========================================
function autoGeminiFeedback(e) {
  if (!e) return; 
  const range = e.range;
  const sheet = range.getSheet();
  
  // ★使い方シートを編集した時はAIが動かないようにブロック
  if (sheet.getName() === "使い方") return;

  if (range.getNumRows() > 1 || range.getNumColumns() > 1) return;

  const col = range.getColumn();
  const row = range.getRow();

  if (col === 5 && row > 1) {
    const textE = range.getValue(); 
    const textD = sheet.getRange(row, 4).getValue(); 
    const textC = sheet.getRange(row, 3).getValue(); 
    
    const cellF = sheet.getRange(row, 6);
    const cellG = sheet.getRange(row, 7);

    if (textE === "") {
      cellF.clearContent();
      cellF.setBackground(null);
      cellG.clearContent();
      cellG.setBackground(null);
      sheet.setRowHeight(row, 21);
      return;
    }
    
    if (textD === "") {
      cellG.setValue("⚠️ 先にD列（5歳児の日本語）を入力してください！");
      return;
    }

    cellF.setValue("🔄 Geminiが回答を作成中...");
    cellG.setValue("🔄 Geminiが添削中...");
    SpreadsheetApp.flush(); 

    const prompt = `あなたはプロの英語教師です。以下のデータをもとにJSON形式で回答してください。\n` +
                   `【データ】\n` +
                   `・元のビジネス日本語: "${textC}"\n` +
                   `・生徒が考えた5歳児レベルの日本語: "${textD}"\n` +
                   `・生徒の英訳: "${textE}"\n\n` +
                   `【出力形式（厳密なJSON形式のみを出力してください。装飾は不要です）】\n` +
                   `{\n` +
                   `  "model_c": "元のビジネス日本語に対する最適なビジネス英語",\n` +
                   `  "model_d": "5歳児レベルの日本語に対する、シンプルで自然な英語",\n` +
                   `  "feedback": "生徒の英訳に対する日本語での優しい添削・アドバイス（文法や単語の選択について）"\n` +
                   `}`;

    const url = "https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-pro:generateContent?key=" + GEMINI_API_KEY;
    
    const payload = {
      "contents": [{"parts": [{"text": prompt}]}],
      "generationConfig": { "responseMimeType": "application/json" }
    };
    
    const options = {
      "method": "post",
      "contentType": "application/json",
      "payload": JSON.stringify(payload),
      "muteHttpExceptions": true
    };

    try {
      const response = UrlFetchApp.fetch(url, options);
      const jsonResponse = JSON.parse(response.getContentText());
      
      if (jsonResponse.error) {
        cellF.setValue("通信エラー");
        cellG.setValue("❌ APIエラー: " + jsonResponse.error.message);
        return;
      }

      let aiText = jsonResponse.candidates[0].content.parts[0].text;
      aiText = aiText.replace(/```json/gi, "").replace(/```/g, "").trim();
      const result = JSON.parse(aiText);

      cellF.setValue(`【ビジネス英訳】\n${result.model_c}\n\n【5歳児日本語の英訳】\n${result.model_d}`);
      cellF.setBackground("#e6fffa");
      cellF.setWrap(true); 

      cellG.setValue(`【Gemini添削】\n${result.feedback}`);
      cellG.setBackground("#fff4e6");
      cellG.setWrap(true);

      sheet.autoResizeRow(row); // 行の高さを自動調整

    } catch (error) {
      cellF.setValue("処理エラー");
      cellG.setValue("❌ エラーが発生しました: " + error.toString());
    }
  }
}

// ==========================================
// 5. メニューの作成
// ==========================================
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('★ビジネス英語ツール')
    .addItem('1. 例文をセットアップ', 'setupSheet')
    .addItem('2. 使い方シートを追加', 'createManualSheet') // ←★追加しました
    .addItem('3. 自動実行の許可設定', 'createTrigger')
    .addToUi();
}