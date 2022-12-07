// import "../../assets/icon-16.png"
// import "../../assets/icon-32.png"
// import "../../assets/icon-80.png"

Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
            document.getElementById('insert').addEventListener('click',writeToSheet)
  }
})

export async function writeToSheet() {
  try {
    await Excel.run(async (context) => {
        let range = context.workbook.getSelectedRange();
        range.load("address");
        await context.sync();
         range.values =[[document.getElementById('output').value]] 
         return context.sync()
    });
  } catch (error) {
    console.error(error);
  }
}
