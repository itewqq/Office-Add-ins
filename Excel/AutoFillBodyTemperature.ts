$("#run").click(() => tryCatch(run));

async function run() {
  await Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    var selectedRange = context.workbook.getSelectedRange();
    selectedRange.load(['address', 'values', 'rowCount','columnCount'])
    return context.sync()
      .then(function () {
        var arr = selectedRange.values;
        for (var i = 0; i < arr.length; ++i) {
          for (var j = 0;j < arr[i].length;++j){
            arr[i][j]=(36.0+0.5*Math.random()).toFixed(1);
          }
        }
        selectedRange.values = arr
      });
  });
}

/** Default helper for invoking an action and handling errors. */
async function tryCatch(callback) {
  try {
    await callback();
  } catch (error) {
    console.error(error);
  }
}
