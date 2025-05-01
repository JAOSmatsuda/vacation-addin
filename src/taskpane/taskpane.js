console.log("taskpane.js ロード: 2025/05/01 17:17");

Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) {
    console.log("📦 Mailbox API バージョンチェック開始");

    if (Office.context.requirements.isSetSupported("Mailbox", "1.7")) {
      console.log("✅ Mailbox 1.7 はサポートされています！");
    } else {
      console.warn("⚠️ Mailbox 1.7 はサポートされていません。");
    }

    document.getElementById("allDayButton").onclick = setAllDayVacation;
    document.getElementById("amHalfDayButton").onclick = setAMVacation;
    document.getElementById("pmHalfDayButton").onclick = setPMVacation;
  }
});

function setAllDayVacation() {
  const item = Office.context.mailbox.item;
  console.log("isAllDayEvent:", item.isAllDayEvent);

  if (!item) {
    console.error("予定アイテムが取得できません");
    return;
  }

  console.log("終日休暇設定開始");

  item.subject.setAsync("終日休暇", (subjectResult) => {
    if (subjectResult.status === Office.AsyncResultStatus.Succeeded) {
      console.log("件名が設定されました");

      if (item.isAllDayEvent && item.isAllDayEvent.setAsync) {
        item.isAllDayEvent.setAsync(true, (allDayResult) => {
          if (allDayResult.status === Office.AsyncResultStatus.Succeeded) {
            console.log("終日イベントが設定されました");
          } else {
            console.error("終日設定に失敗:", allDayResult.error.message);
          }
        });
      } else {
        console.warn("isAllDayEvent はサポートされていません");
      }

    } else {
      console.error("件名の設定に失敗:", subjectResult.error.message);
    }
  });
}

function setAMVacation() {
  const item = Office.context.mailbox.item;
  if (!item) return;

  item.subject.setAsync("AM半休", (subjectResult) => {
    if (subjectResult.status !== Office.AsyncResultStatus.Succeeded) {
      console.error("件名の設定に失敗:", subjectResult.error.message);
      return;
    }

    const start = new Date();
    start.setHours(9, 0, 0, 0);
    const end = new Date();
    end.setHours(13, 0, 0, 0);

    item.start.setAsync(start, (startResult) => {
      if (startResult.status !== Office.AsyncResultStatus.Succeeded) {
        console.error("開始時刻エラー:", startResult.error.message);
        return;
      }

      item.end.setAsync(end, (endResult) => {
        if (endResult.status !== Office.AsyncResultStatus.Succeeded) {
          console.error("終了時刻エラー:", endResult.error.message);
        }
      });
    });
  });
}


function setPMVacation() {
  const item = Office.context.mailbox.item;
  if (!item) return;

  item.subject.setAsync("PM半休", (subjectResult) => {
    if (subjectResult.status !== Office.AsyncResultStatus.Succeeded) {
      console.error("件名の設定に失敗:", subjectResult.error.message);
      return;
    }

    const start = new Date();
    start.setHours(14, 0, 0, 0);
    const end = new Date();
    end.setHours(18, 15, 0, 0);

    item.start.setAsync(start, (startResult) => {
      if (startResult.status !== Office.AsyncResultStatus.Succeeded) {
        console.error("開始時刻エラー:", startResult.error.message);
        return;
      }

      item.end.setAsync(end, (endResult) => {
        if (endResult.status !== Office.AsyncResultStatus.Succeeded) {
          console.error("終了時刻エラー:", endResult.error.message);
        }
      });
    });
  });
}