(function () {
  const oWorksheet = Api.GetActiveSheet();
  const query = oWorksheet.GetRange("A1").GetValue();

  //api url. Consumes available requests
  const url = `http://localhost:3000/search?query=${query}`;

  //testUrl. Fetches data from the sampleData.js file from the server. Doesnt consume requests from the API.
  const testUrl = `http://localhost:3000/test?query=${query}`;

  const data = null;
  const xhr = new XMLHttpRequest();

  xhr.onload = () => {
    const oData = JSON.parse(xhr.response);
    const { organic_results: results } = oData;

    //take out titles and respective links in array
    const title = results.map((i) => i.title);
    const link = results.map((i) => i.link);

    //insert titles
    let nRowTitle = 2;
    for (let i = 0; i < title.length; i++) {
      const text = JSON.stringify(title[i]);
      oWorksheet.GetRangeByNumber(nRowTitle, 0).SetValue(text);
      oWorksheet.GetRangeByNumber(nRowTitle, 0).AutoFit(false, true);
      nRowTitle++;
    }

    //insert links parallel to the titles
    let nRowLink = 2;
    for (let i = 0; i < link.length; i++) {
      const text = JSON.stringify(link[i]);
      oWorksheet.GetRangeByNumber(nRowLink, 1).SetValue(text);
      oWorksheet.GetRangeByNumber(nRowLink, 1).AutoFit(false, true);
      nRowLink++;
    }
  };

  function reloadCellValues() {
    let reload = setTimeout(function () {
      Api.asc_calculate(Asc.c_oAscCalculateType.All);
    }, 5000);
  }

  reloadCellValues();
  xhr.open("GET", testUrl);
  xhr.send(data);
})();
