const scriptVersion = "2025/05/02 10:55";

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

// Promise化ユーティリティ
function setAsyncWrapper(method, value) {
  return new Promise((resolve, reject) => {
    method.setAsync(value, (result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        resolve();
      } else {
        reject(result.error.message);
      }
    });
  });
}

async function setAllDayVacation() {
  const item = Office.context.mailbox.item;
  const isClassic = Office.context.mailbox.diagnostics.hostName === "Outlook";
  console.log("終日休暇設定開始");

  if (!item) {
    console.error("予定アイテムが取得できません");
    showMessage("予定の取得に失敗しました", true);
    return;
  }

  try {
    await setAsyncWrapper(item.subject, "終日休暇");
    console.log("件名が設定されました");

    await setAsyncWrapper(item.busyStatus, Office.MailboxEnums.BusyStatus.OOF);
    console.log("公開方法が不在に設定されました");

    if (isClassic && item.isAllDayEvent && typeof item.isAllDayEvent.setAsync === "function") {
      await setAsyncWrapper(item.isAllDayEvent, true);
      console.log("Classic: 終日設定が成功しました");
    } else {
      const start = new Date();
      start.setHours(0, 0, 0, 0);
      const end = new Date();
      end.setHours(23, 59, 0, 0);

      await setAsyncWrapper(item.start, start);
      await setAsyncWrapper(item.end, end);
      console.log("New: 終日風スケジュール設定成功");
    }

    showMessage("終日に設定されました");
  } catch (error) {
    console.error("終日休暇設定エラー:", error);
    showMessage("終日設定に失敗しました", true);
  }
}

async function setHalfDay(startHour, endHour, endMinute, subjectText) {
  const item = Office.context.mailbox.item;
  console.log(`${subjectText}設定開始`);

  if (!item) {
    console.error("予定アイテムが取得できません");
    showMessage("予定の取得に失敗しました", true);
    return;
  }

  try {
    await setAsyncWrapper(item.subject, subjectText);
    console.log("件名が設定されました");

    await setAsyncWrapper(item.busyStatus, Office.MailboxEnums.BusyStatus.OOF);
    console.log("公開方法が不在に設定されました");

    if (item.isAllDayEvent && typeof item.isAllDayEvent.setAsync === "function") {
      await setAsyncWrapper(item.isAllDayEvent, false);
      console.log("終日解除成功");
    }

    const start = new Date();
    start.setHours(startHour, 0, 0, 0);
    const end = new Date();
    end.setHours(endHour, endMinute, 0, 0);

    await setAsyncWrapper(item.start, start);
    await setAsyncWrapper(item.end, end);

    console.log(`${subjectText}の時間が設定されました`);
    showMessage(`${subjectText}として設定されました`);
  } catch (error) {
    console.error(`${subjectText}設定エラー:`, error);
    showMessage(`${subjectText}設定に失敗しました`, true);
  }
}

Office.onReady(() => {
  updateVersionDisplay();

  document.getElementById("allDayButton").onclick = setAllDayVacation;
  document.getElementById("amHalfDayButton").onclick = () => setHalfDay(9, 14, 0, "AM半休");
  document.getElementById("pmHalfDayButton").onclick = () => setHalfDay(14, 18, 15, "PM半休");
});