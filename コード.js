// 実際の拡張子は ".gs"
// "コード.gs"

// LINE developersのメッセージ送受信設定に記載のアクセストークン
const LINE_TOKEN = "個別のアクセストークン"
const LINE_URL = "https://api.line.me/v2/bot/message/reply";


//postリクエストを受取ったときに発火する関数
function doPost(e) {
  function addMsg(msg) {
    const add = {
      type: "text",
      text: msg,
    };
    messages.push(add);
  }

  function showDate(year, month, day) {
    return year + "年" + month + "月" + day + "日";
  }

  // 応答用Tokenを取得
  const replyToken = JSON.parse(e.postData.contents).events[0].replyToken;
  // メッセージを取得
  const userMessage = JSON.parse(e.postData.contents).events[0].message.text;
  // ユーザーIDを取得
  const userid = JSON.parse(e.postData.contents).events[0].source.userId;

  //メッセージを改行ごとに分割
  const all_msg = userMessage.split(",");
  const msg_num = all_msg.length;

  const messages = [];

  if (
    all_msg.length == 4 &&
    all_msg[1] >= 2020 &&
    all_msg[2] >= 0 &&
    all_msg[2] <= 12 &&
    all_msg[3] >= 0 &&
    all_msg[2] <= 31
  ) {
    // 日付のバリデーションが必要
    // ***************************
    // スプレットシートからデータを抽出
    // ***************************
    // 1. 今開いている（紐付いている）スプレッドシートを定義
    const sheet = SpreadsheetApp.getActiveSpreadsheet();
    // 2. ここでは、デフォルトの「シート1」の名前が書かれているシートを呼び出し
    const listSheet = sheet.getSheetByName("シート1");
    // 3. 最終列の列番号を取得
    const numColumn = listSheet.getLastColumn();
    // 4. 最終行の行番号を取得
    const numRow = listSheet.getLastRow() - 1;
    // 5. 範囲を指定（上、左、右、下）
    const topRange = listSheet.getRange(1, 1, 1, numColumn); // 一番上のオレンジ色の部分の範囲を指定
    const dataRange = listSheet.getRange(2, 1, numRow, numColumn); // データの部分の範囲を指定
    // 6. 値を取得
    //const topData = topRange.getValues(); // 一番上のオレンジ色の部分の範囲の値を取得
    const data = dataRange.getValues(); // データの部分の範囲の値を取得
    const dataNum = data.length + 2; // 新しくデータを入れたいセルの列の番号を取得

    const plans = data.filter((value) => {
      // useridに紐付いた予定の抽出
      if (value.indexOf(userid) !== -1) {
        return value;
      }
    });
    if (plans.length < 5) {
      // ***************************
      // スプレッドシートにデータを入力
      // ***************************
      // 最終列の番号まで、順番にスプレッドシートの左からデータを新しく入力
      for (let i = 0; i < msg_num + 1; i++) {
        if (i == 0) {
          SpreadsheetApp.getActiveSheet()
            .getRange(dataNum, i + 1)
            .setValue(userid);
        } else {
          SpreadsheetApp.getActiveSheet()
            .getRange(dataNum, i + 1)
            .setValue(all_msg[i - 1]);
        }
      }
      addMsg(
        "「" +
          all_msg[0] +
          "」を " +
          showDate(all_msg[1], all_msg[2], all_msg[3]) +
          " で登録しました。"
      );
    } else {
      addMsg("予定は5つまでしか登録できません。");
    }
  } else if (all_msg[0] == "削除" && all_msg[1] !== null) {
    const sh = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = sh.getSheetByName("シート1");
    const last_row = sheet.getLastRow();
    let flag = 0;
    for (let i = 2; i <= last_row; i++) {
      const ran = sheet.getRange("A" + i);
      const val = ran.getDisplayValue();
      const range = sheet.getRange("B" + i);
      const value = range.getDisplayValue();
      if (val == userid && value == all_msg[1]) {
        sheet.deleteRow(i);
        addMsg("予定「" + all_msg[1] + "」を削除しました。");
        i = i - 1;
        flag = 1;
      }
    }
    if (flag == 0) {
      addMsg("削除する予定がありません。");
    }
  } else if (all_msg[0] == "ヘルプ") {
    addMsg("毎日0時になると 予定までの残りの日数を通知します。");
    addMsg(
      "【登録】\n「タイトル,年,月,日」と入力して予定を登録してください。\n例）ライブ,2020,12,25"
    );
    addMsg("【一覧】\n「一覧」と入力してください。予定を確認できます。");
    addMsg(
      "【削除】\n「削除,予定タイトル」と入力してください。予定を削除できます。\n例）削除,ライブ"
    );
    addMsg("予定は5つまでしか登録できません。\n当日になると、その予定は自動的に削除されます。");
  } else if (all_msg[0] == "一覧") {
    // 1. 今開いている（紐付いている）スプレッドシートを定義
    const sheet = SpreadsheetApp.getActiveSpreadsheet();
    // 2. ここでは、デフォルトの「シート1」の名前が書かれているシートを呼び出し
    const listSheet = sheet.getSheetByName("シート1");

    // 3. 最終列の列番号を取得
    const numColumn = listSheet.getLastColumn();
    // 4. 最終行の行番号を取得
    const numRow = listSheet.getLastRow() - 1;
    // 5. 範囲を指定（上、左、右、下）
    const topRange = listSheet.getRange(1, 1, 1, numColumn); // 一番上のオレンジ色の部分の範囲を指定
    const dataRange = listSheet.getRange(2, 1, numRow, numColumn); // データの部分の範囲を指定

    // 6. 値を取得
    const topData = topRange.getValues(); // 一番上のオレンジ色の部分の範囲の値を取得
    const data = dataRange.getValues(); // データの部分の範囲の値を取得
    const dataNum = data.length + 2; // 新しくデータを入れたいセルの列の番号を取得

    const plans = data.filter((value) => {
      // useridに紐付いた予定の抽出
      if (value.indexOf(userid) !== -1) {
        return value;
      }
    });

    if (plans == null) {
      addMsg("予定はありません。");
    } else {
      for (let i in plans) {
        addMsg(
          "「" +
            plans[i][1] +
            "」" +
            showDate(plans[i][2], plans[i][3], plans[i][4]) + 
            "\nあと " +
            last_day(plans[i][2], plans[i][3], plans[i][4]) +
            "日"
        );
      }
    }
  } else {
    addMsg("不正な入力です。「ヘルプ」でコマンドを確認してください。");
  }

  //lineで返答する
  UrlFetchApp.fetch(LINE_URL, {
    headers: {
      "Content-Type": "application/json; charset=UTF-8",
      Authorization: `Bearer ${LINE_TOKEN}`,
    },
    method: "post",
    payload: JSON.stringify({
      replyToken: replyToken,
      messages: messages,
    }),
  });

  ContentService.createTextOutput(
    JSON.stringify({ content: "post ok" })
  ).setMimeType(ContentService.MimeType.JSON);
}

//通知機能
function push_plans() {
  // 1. 今開いている（紐付いている）スプレッドシートを定義
  const sheet = SpreadsheetApp.getActiveSpreadsheet();
  // 2. ここでは、デフォルトの「シート1」の名前が書かれているシートを呼び出し
  const listSheet = sheet.getSheetByName("シート1");

  // 3. 最終列の列番号を取得
  const numColumn = listSheet.getLastColumn();
  // 4. 最終行の行番号を取得
  const numRow = listSheet.getLastRow();
  // 5. 範囲を指定（上、左、右、下）
  const topRange = listSheet.getRange(1, 1, 1, numColumn); // 一番上のオレンジ色の部分の範囲を指定
  const dataRange = listSheet.getRange(2, 1, numRow, numColumn); // データの部分の範囲を指定
  // 6. 値を取得
  const data = dataRange.getValues(); // データの部分の範囲の値を取得

  // ユーザーの名前だけ取得
  const temp = JSON.parse(JSON.stringify(data));
  const exe = [1, 2, 3, 4]; //削除したい位置（先頭が0であることに注意）
  for (let i = 0; i < temp.length; i++) {
    //このfor文で行を回す
    for (let j = 0; j < exe.length; j++) {
      temp[i].splice(exe[j] - j, 1);
    }
  }
  const temp2 = temp.reduce((pre, current) => {
    pre.push(...current);
    return pre;
  }, []);
  const users = [...new Set(temp2)];
  users.pop();
  const boxs = {};

  users.forEach((user, i) => {
    boxs["user" + i] = data.filter((value) => {
      if (value.indexOf(user) !== -1) {
        return value;
      }
    });
  });

  users.forEach((user, num) => {
    var postData = {
      to: user,
      messages: [],
    };

    for (let i = 0; i < Math.min(boxs["user" + num].length, 5); i++) {
      if (
        last_day(
          boxs["user" + num][i][2],
          boxs["user" + num][i][3],
          boxs["user" + num][i][4]
        ) > 0
      ) {
        const add = {
          type: "text",
          text:
            "「" +
            boxs["user" + num][i][1] +
            "」まで あと " +
            last_day(
              boxs["user" + num][i][2],
              boxs["user" + num][i][3],
              boxs["user" + num][i][4]
            ) +
            "日です。",
        };
        postData.messages.push(add);
      } else if (
        last_day(
          boxs["user" + num][i][2],
          boxs["user" + num][i][3],
          boxs["user" + num][i][4]
        ) == 0
      ) {
        const add = {
          type: "text",
          text: "今日は「" + boxs["user" + num][i][1] + "」の日です。",
        };
        postData.messages.push(add);
      }
    }

    var headers = {
      "Content-Type": "application/json",
      Authorization: "Bearer " + LINE_TOKEN,
    };

    var options = {
      method: "post",
      headers: headers,
      payload: JSON.stringify(postData),
    };

    var response = UrlFetchApp.fetch(
      "https://api.line.me/v2/bot/message/push",
      options
    );
  });
}
// 残り日数 func
function last_day(year, month, day) {
  let today = new Date();
  today = new Date(today.getFullYear(), today.getMonth(), today.getDate());
  let planday = new Date(year, month - 1, day);
  return (planday - today) / 86400000;
}

// 今日、今日より前の予定の削除
function delete_expired_plans() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet();
  const sh = sheet.getSheetByName("シート1");
  const last_row = sh.getLastRow();
  let flag = 2;
  let i = 2;
  do {
    const year_ran = sh.getRange("C" + i);
    const year = year_ran.getDisplayValue();
    const month_ran = sh.getRange("D" + i);
    const month = month_ran.getDisplayValue();
    const day_ran = sh.getRange("E" + i);
    const day = day_ran.getDisplayValue();
    if (last_day(year, month, day) <= 0) {
      sh.deleteRow(i);
      i = i - 1;
    }
    flag++;
    i++;
  } while (flag <= last_row);
}
