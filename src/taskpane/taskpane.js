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
function setAMVacation() {
  const item = Office.context.mailbox.item;
  item.subject.setAsync("AM半休", () => {
    item.isAllDayEvent.setAsync(false, () => {
      const start = new Date();
      start.setHours(9, 0, 0, 0);
      const end = new Date();
      end.setHours(13, 0, 0, 0);

      item.start.setAsync(start, () => {
        item.end.setAsync(end, () => {
          console.log("AM半休の時間が設定されました。");
        });
      });
    });
  });
}

// PM半休の設定
function setPMVacation() {
  const item = Office.context.mailbox.item;
  item.subject.setAsync("PM半休", () => {
    item.isAllDayEvent.setAsync(false, () => {
      const start = new Date();
      start.setHours(14, 0, 0, 0);
      const end = new Date();
      end.setHours(18, 15, 0, 0);

      item.start.setAsync(start, () => {
        item.end.setAsync(end, () => {
          console.log("PM半休の時間が設定されました。");
        });
      });
    });
  });
}