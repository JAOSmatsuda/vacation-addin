console.log("taskpane.js ロード: 2025/05/02 09:21");

function showMessage(message, isError = false) {
  const messageDiv = document.getElementById("message");
  if (!messageDiv) return;

  messageDiv.textContent = message;
  messageDiv.style.color = isError ? "red" : "green";
  messageDiv.style.display = "block";

  setTimeout(() => {
    messageDiv.style.display = "none";
  }, 5000);
}

function fallbackAllDaySet(item) {
  const start = new Date();
  start.setHours(0, 0, 0, 0);
  const end = new Date(start);
  end.setDate(start.getDate() + 1);

  item.start.setAsync(start, (startResult) => {
    if (startResult.status !== Office.AsyncResultStatus.Succeeded) {
      console.error("開始時刻設定エラー:", startResult.error.message);
      showMessage("開始時刻の設定に失敗しました", true);
      return;
    }

    item.end.setAsync(end, (endResult) => {
      if (endResult.status !== Office.AsyncResultStatus.Succeeded) {
        console.error("終了時刻設定エラー:", endResult.error.message);
        showMessage("終了時刻の設定に失敗しました", true);
      } else {
        console.log("New: 終日風スケジュール設定成功");
        showMessage("終日に設定されました");
      }
    });
  });
}

function setAllDayVacation() {
  const item = Office.context.mailbox.item;
  console.log("終日休暇設定開始");

  if (!item) {
    console.error("予定アイテムが取得できません");
    showMessage("予定の取得に失敗しました", true);
    return;
  }

  item.subject.setAsync("終日休暇", () => {
    console.log("件名が設定されました");
  });

  if (Office.context.requirements.isSetSupported("Mailbox", "1.7") &&
      item.isAllDayEvent &&
      typeof item.isAllDayEvent.setAsync === "function") {

    item.isAllDayEvent.setAsync(true, (result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        console.log("Classic: 終日設定成功");
        showMessage("終日に設定されました");
      } else {
        console.warn("Classic: isAllDayEvent 設定失敗、fallback へ");
        fallbackAllDaySet(item);
      }
    });

  } else {
    console.log("New Outlook: isAllDayEvent 非対応。fallback 使用");
    fallbackAllDaySet(item);
  }
}

function setHalfDay(startHour, endHour, subjectText) {
  const item = Office.context.mailbox.item;
  console.log(`${subjectText}設定開始`);

  if (!item) {
    console.error("予定アイテムが取得できません");
    showMessage("予定の取得に失敗しました", true);
    return;
  }

  item.subject.setAsync(subjectText, () => {
    console.log("件名が設定されました");
  });

  if (item.isAllDayEvent && typeof item.isAllDayEvent.setAsync === "function") {
    item.isAllDayEvent.setAsync(false, (result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        console.log("終日解除成功");
      } else {
        console.error("終日解除に失敗:", result.error.message);
      }
    });
  }

  const start = new Date();
  start.setHours(startHour, 0, 0, 0);
  const end = new Date();
  end.setHours(endHour, 0, 0, 0);

  item.start.setAsync(start, (startResult) => {
    if (startResult.status !== Office.AsyncResultStatus.Succeeded) {
      console.error("開始時刻設定エラー:", startResult.error.message);
      showMessage("開始時刻の設定に失敗しました", true);
      return;
    }

    item.end.setAsync(end, (endResult) => {
      if (endResult.status !== Office.AsyncResultStatus.Succeeded) {
        console.error("終了時刻設定エラー:", endResult.error.message);
        showMessage("終了時刻の設定に失敗しました", true);
      } else {
        console.log(`${subjectText}の時間が設定されました`);
        showMessage(`${subjectText}として設定されました`);
      }
    });
  });
}

Office.onReady(() => {
  document.getElementById("allDayButton").onclick = setAllDayVacation;
  document.getElementById("amHalfDayButton").onclick = () => setHalfDay(9, 13, "AM半休");
  document.getElementById("pmHalfDayButton").onclick = () => setHalfDay(13, 17, "PM半休");
});