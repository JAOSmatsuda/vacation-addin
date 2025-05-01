console.log("taskpane.js ロード: 2025/05/01 17:44");

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
  if (!item) {
    console.error("予定アイテムが取得できません");
    showMessage("予定アイテムが取得できませんでした。");
    return;
  }

  console.log("終日休暇設定開始");

  // 件名を設定
  item.subject.setAsync("終日休暇", (subjectResult) => {
    if (subjectResult.status !== Office.AsyncResultStatus.Succeeded) {
      console.error("件名の設定に失敗:", subjectResult.error.message);
      showMessage("件名の設定に失敗しました。");
      return;
    }

    console.log("件名が設定されました");

    // isAllDayEvent が利用可能か確認
    if (item.isAllDayEvent && typeof item.isAllDayEvent.setAsync === "function") {
      item.isAllDayEvent.setAsync(true, (allDayResult) => {
        if (allDayResult.status === Office.AsyncResultStatus.Succeeded) {
          console.log("終日イベントが設定されました");
          showMessage("『終日休暇』を設定しました。");
        } else {
          console.error("終日イベントの設定に失敗:", allDayResult.error.message);
          showMessage("件名は設定されましたが、終日設定に失敗しました。");
        }
      });
    } else {
      // New Outlook など isAllDayEvent 未対応の場合
      console.warn("isAllDayEvent が未対応のため、手動対応を促します。");
      showMessage("件名を『終日休暇』に設定しました。終日のチェックは手動で ON にしてください。");
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

function showMessage(text, duration = 5000) {
  const messageDiv = document.getElementById("message");
  if (!messageDiv) return;

  messageDiv.textContent = text;
  messageDiv.style.display = "block";

  setTimeout(() => {
    messageDiv.textContent = "";
    messageDiv.style.display = "none";
  }, duration);
}