/* global Office */

Office.onReady(() => {
  // Office.jsが読み込み完了
});

/**
 * 件名を「終日休暇」に設定し、終日フラグをONにする
 * @param {Office.AddinCommands.Event} event
 */
function setVacationEvent(event) {
  const item = Office.context.mailbox.item;

  item.subject.setAsync("終日休暇", (subjectResult) => {
    if (subjectResult.status !== Office.AsyncResultStatus.Succeeded) {
      console.error("件名設定失敗:", subjectResult.error.message);
    }
  });

  item.isAllDayEvent.setAsync(true, (allDayResult) => {
    if (allDayResult.status !== Office.AsyncResultStatus.Succeeded) {
      console.error("終日設定失敗:", allDayResult.error.message);
    }
  });

  event.completed();
}

Office.actions.associate("setVacationEvent", setVacationEvent);