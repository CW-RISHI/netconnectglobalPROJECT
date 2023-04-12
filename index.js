const data1 = document.getElementById("data1");
const table1 = document.getElementById("table1");
const table = document.getElementById("table");
const selection = document.getElementById("selection");


async function excelSheetJsonData () {

  const excelData = await fetch("./excel.xlsx");
  const excelDataArray = await excelData.arrayBuffer();
  const workBook = XLSX.read(excelDataArray);


  const workSheet = workBook.SheetNames;
  const data = workSheet.map((name) => {
    let html = XLSX.utils.sheet_to_json(workBook.Sheets[name]);
    return html
  });

  return data;
};


const changeHandler = (e) => {
  if(e.target.value === "Data 1") {
    excelSheetJsonData().then(res => {
      const xArray = res[0].map(item => item.Country);
      const yArray = res[0].map(item => item["Sum of Sales (€million)"]);

      const data = [{
        x: xArray,
        y: yArray,
        marker: {color: '#78ADD2'},
        mode: "lines",
        type: "scatter"
      }];

      const layout = {
        xaxis: {range: [], title: "Country" },
        yaxis: {range: [0, 50000000], title: "Sum of Sales (€million)"},
        title: "Anually sum of sales of countries in million euro"
      }
      Plotly.newPlot(data1, data, layout);
      data1.style.display = "block";
      table1.style.display = "none";
    });
  } else if (e.target.value === "Data 2") {
    excelSheetJsonData().then(res => {
      const xArray = res[1].map(item => item.Year);
      const yArray = res[1].map(item => item["Sum of Sales (€million)"]);

      const data = [{
        x: xArray,
        y: yArray,
        marker: {color: '#0072AA'},
        mode: "lines",
        type: "scatter"
      }];

      const layout = {
        xaxis: {range: [2010, 2020], title: "Years" },
        yaxis: {range: [0, 50000000], title: "Sum of Sales (€million)"},
        title: "Anually sum of sales in million euro"
      }
      Plotly.newPlot(data1, data, layout);
      data1.style.display = "block";
      table1.style.display = "none";
    });
  } else if (e.target.value === "Data 3") {
    excelSheetJsonData().then(res => {
      const xArray = res[2].map(item => item["Industrial sector"]);
      const yArray = res[2].map(item => item["Sum of Profits (€million)"]);

      const data = [{
        x: xArray,
        y: yArray,
        marker: {color: '#A4D0A0'},
        type: "bar",
        mode: "bar"
      }];

      const layout = {
        xaxis: {range: [0, 50], title: "Countries" },
        yaxis: {range: [0, 1000000], title: "Sum of Profits (€million)"},
        title: "Annual profit across different sector"
      }
      Plotly.newPlot(data1, data, layout);
      data1.style.display = "block";
      table1.style.display = "none";
    });
  } else if (e.target.value === "Data 4") {
    excelSheetJsonData().then(res => {
      const data = res[3];
      for(let i = 0; i < res[3].length; i++) {
        let row = `
            <tr class = "tr1">
              <td>${data[i].Year}</td>
              <td>${data[i]["Sum of Sales (€million)"]}</td>
              <td>${data[i]["Sum of Capex (€million)"]}</td>
              <td>${data[i]["Sum of Profits (€million)"]}</td>
              <td>${data[i]["Sum of Market cap (€million)"]}</td>
            </tr>
        `
        data1.style.display = "none";
        table1.style.display = "block";
        table.innerHTML += row
      }
    });
  } else {
    return;
  }
};

selection.addEventListener("change", changeHandler)

