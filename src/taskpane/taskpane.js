// グローバル変数にボタンを格納
const allDayButton = document.getElementById("allDayButton");
const amHalfDayButton = document.getElementById("amHalfDayButton");
const pmHalfDayButton = document.getElementById("pmHalfDayButton");

Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) {
    // ボタンのクリックイベントを設定
    allDayButton.onclick = setAllDayVacation;
    amHalfDayButton.onclick = setAMHalfDay;
    pmHalfDayButton.onclick = setPMHalfDay;
  }
});

function setAllDayVacation() {
  const item = Office.context.mailbox.item;

  item.subject.setAsync("終日休暇", (subjectResult) => {
    if (subjectResult.status === Office.AsyncResultStatus.Succeeded) {
      console.log("件名が設定されました。");

      if (item.isAllDayEvent) {
        item.isAllDayEvent.setAsync(true, (allDayResult) => {
          if (allDayResult.status === Office.AsyncResultStatus.Succeeded) {
            console.log("終日イベントが設定されました。");
          } else {
            console.error("終日設定に失敗:", allDayResult.error.message);
          }
        });
      } else {
        console.warn("isAllDayEvent はこの環境ではサポートされていません。");
      }

    } else {
      console.error("件名の設定に失敗:", subjectResult.error.message);
    }
  });
}

// AM半休の設定
function setAMHalfDay() {
  const item = Office.context.mailbox.item;
  item.subject.setAsync("AM半休", () => {
    item.isAllDayEvent.setAsync(false); // 終日イベントをオフ
    item.start.setAsync(new Date(item.start.date.getFullYear(), item.start.date.getMonth(), item.start.date.getDate(), 9, 0)); // 開始時刻をAM9:00に設定
    item.end.setAsync(new Date(item.start.date.getFullYear(), item.start.date.getMonth(), item.start.date.getDate(), 12, 0)); // 終了時刻をPM12:00に設定
  });
}

// PM半休の設定
function setPMHalfDay() {
  const item = Office.context.mailbox.item;
  item.subject.setAsync("PM半休", () => {
    item.isAllDayEvent.setAsync(false); // 終日イベントをオフ
    item.start.setAsync(new Date(item.start.date.getFullYear(), item.start.date.getMonth(), item.start.date.getDate(), 14, 0)); // 開始時刻をPM2:00に設定
    item.end.setAsync(new Date(item.start.date.getFullYear(), item.start.date.getMonth(), item.start.date.getDate(), 18, 15)); // 終了時刻をPM6:15に設定
  });
}