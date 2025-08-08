Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
  }
});

export async function openSettings() {
    await Excel.run(async (context) => {
        console.log("Opening settings dialog");
        Office.context.ui.displayDialogAsync(
            'https://everett022.github.io/COR-AUTO/orderingSet.html',
            {height: 65, width: 85},
        );
        await context.sync();
    });
}
