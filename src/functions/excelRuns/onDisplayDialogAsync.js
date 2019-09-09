export const onDisplayDialogAsync = async (context) => {
  window.Office.context.ui.displayDialogAsync('https://baidu.com', {
      promptBeforeOpen: false,
      height: 30,
      width: 20
    },
    function (asyncResult) {
      console.log(asyncResult)
      // dialog = asyncResult.value;
      // dialog.addEventHandler(Office.EventType.DialogMessageReceived, processMessage);
    }
  )
  await context.sync()
    .then(() => {
    })
    .catch(error => {
      console.log(`Error: ${error}`)
    })
}