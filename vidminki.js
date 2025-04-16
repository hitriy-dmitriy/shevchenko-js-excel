// in excel install script-lab add-ins
// import library shevchenko js
// <script src="https://cdn.jsdelivr.net/npm/shevchenko@3.1.4/dist/umd/shevchenko.min.js"></script>


document.getElementById("run").addEventListener("click", () => tryCatch(run));

async function run() {
  await Excel.run(async (context) => {
    const sheet1 = context.workbook.worksheets.getItem("Аркуш1"); //назва листа звідки брати
    const sheet2 = context.workbook.worksheets.getItem("Аркуш2");//назва листа куди ставити
    
    //початок
    const cellStart = "A1";
    //кінець 
    const cellEnd = "A112"

    const rangeCell = sheet1.getRange(`${cellStart}:${cellEnd}`);
    rangeCell.load("values");

    await context.sync();

    const rows = rangeCell.values.flat();
    const tmp = rows.map((x) => x.trim().split(" "));
    let antroponim1 = {
      givenName: "",
      patronymicName: "",
      familyName: ""
    };
    for (let i = 0; i < tmp.length; i++) {
      //console.log(tmp[i], tmp[i][0])
      antroponim1 = {
        givenName: tmp[i][1],
        patronymicName: tmp[i][2],
        familyName: tmp[i][0]
      };
      const result = await vidminok(antroponim1);
      await context.sync();
      //console.log(result)

      sheet2.getRange(`C${i + 1}`).values = [[result]];
      await context.sync();
    }

    async function vidminok(anthroponym) {
      const gender = await shevchenko.detectGender(anthroponym);
      if (gender == null) {
        throw new Error("Failed to detect grammatical gender.");
      }

      const input = { ...anthroponym, gender };
      
      // Давальний відмінок
      const output = await shevchenko.inDative(input);

      return `${output.familyName} ${output.givenName} ${output.patronymicName}`;
    }

    /*  const gender = await shevchenko.detectGender(anthroponym); // "feminine"
    if (gender == null) {
      throw new Error("Failed to detect grammatical gender.");
    }

    const input = { ...anthroponym, gender };

    const output = await shevchenko.inVocative(input);

    console.log(output); */

    await context.sync();
  });
}

/** Default helper for invoking an action and handling errors. */
async function tryCatch(callback) {
  try {
    await callback();
  } catch (error) {
    // Note: In a production add-in, you'd want to notify the user through your add-in's UI.
    console.error(error);
  }
}
