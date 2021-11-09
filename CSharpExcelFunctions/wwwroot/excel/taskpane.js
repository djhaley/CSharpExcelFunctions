window.functionsRef = null;
window.setFunctionsRef = fnRef => window.functionsRef = fnRef;

Office.initialize = () => {
  setTimeout(() => {
      document.getElementById("sideload-msg").style.display = "none";
      document.getElementById("app-body").style.display = "flex";
      document.getElementById("run").onclick = run;
      console.log(document.getElementById("sideload-msg"));
  }, 100);
};

function run() {
  try {
    Excel.run(async context => {
      /**
       * Insert your Excel code here
       */
      const range = context.workbook.getSelectedRange();

        let x = 0;
        x = await add(1, 4);

        console.log(x);

      // Read the range address
      range.load("address");

      // Update the fill color
      range.format.fill.color = "yellow";
      range.values = [[ x ]];

      await context.sync();
      console.log(`The range address was ${range.address}.`);
    });
  } catch (error) {
    console.error(error);
  }
}
