const scriptVersion = "2025/05/02 14:52";

function updateVersionDisplay() {
  const versionElement = document.getElementById("version");
  if (versionElement) {
    versionElement.textContent = `Version: ${scriptVersion}`;
  }
}

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

const OfficeBusyStatus = {
  Free: 0,
  Tentative: 1,
  Busy: 2,
  OOF: 3,
  WorkingElsewhere: 4,
  NoData: 5
};

function fallbackAllDaySet(item) {
  const start = new Date();
  start.setHours(0, 0, 0, 0);
  const end = new Date(start);
  end.setDate(start.getDate() + 1);

  item.start.setAsync(start, () => {
    item.end.setAsync(end, () => {
      showMessage("終日に設定されました");
    });
  });
}

function setAllDayVacation() {
  const item = Office.context.mailbox.item;
  const isClassic = Office.context.mailbox.diagnostics.hostName === "Outlook";
  console.log("終日休暇設定開始");

  if (!item) {
    console.error("予定アイテムが取得できません");
    showMessage("予定の取得に失敗しました", true);
    return;
  }

  item.subject.setAsync("終日休暇", () => {
    console.log("件名が設定されました");
  });

  if (item.busyStatus && typeof item.busyStatus.setAsync === "function") {
    item.busyStatus.setAsync(OfficeBusyStatus.OOF, (result) => {
      if (result.status !== Office.AsyncResultStatus.Succeeded) {
        console.error("公開方法の設定に失敗:", result.error.message);
        showMessage("公開方法の設定に失敗しました", true);
      } else {
        console.log("公開方法が外出中に設定されました");
        showMessage("公開方法を「外出中」に設定しました");
      }
    });
  }

  if (isClassic && item.isAllDayEvent && typeof item.isAllDayEvent.setAsync === "function") {
    item.isAllDayEvent.setAsync(true, (result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        console.log("Classic: 終日設定が成功しました");
        showMessage("終日に設定されました");
      } else {
        console.error("Classic: 終日設定に失敗:", result.error.message);
        showMessage("終日設定に失敗しました", true);
      }
    });
  } else {
    const start = new Date();
    start.setHours(0, 0, 0, 0);
    const end = new Date();
    end.setHours(23, 59, 0, 0);

    item.start.setAsync(start, () => {
      item.end.setAsync(end, () => {
        showMessage("終日に設定されました");
      });
    });
  }
}

function setHalfDay(startHour, endHour, endMinute, subjectText) {
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

  if (item.busyStatus && typeof item.busyStatus.setAsync === "function") {
    item.busyStatus.setAsync(OfficeBusyStatus.OOF, (result) => {
      if (result.status !== Office.AsyncResultStatus.Succeeded) {
        console.error("公開方法の設定に失敗:", result.error.message);
        showMessage("公開方法の設定に失敗しました", true);
      } else {
        console.log("公開方法が外出中に設定されました");
        showMessage("公開方法を「外出中」に設定しました");
      }
    });
  }

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
  end.setHours(endHour, endMinute, 0, 0);

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
  updateVersionDisplay();

  document.getElementById("allDayButton").onclick = setAllDayVacation;
  document.getElementById("amHalfDayButton").onclick = () => setHalfDay(9, 14, 0, "AM半休");
  document.getElementById("pmHalfDayButton").onclick = () => setHalfDay(14, 18, 15, "PM半休");
});