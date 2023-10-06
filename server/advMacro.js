(function () {
  const oWorksheet = Api.GetActiveSheet();

  let query = [];
  for (let i = 2; i < 5; i++) {
    const value = String(oWorksheet.GetRange("A" + i).GetValue());
    if (value !== null && value.length !== 0) {
      query.push(value);
    }
  }

  function populate(count, title, link) {
    let nRowTitle = 9;
    for (let j = 0; j < title.length; j++) {
      let passedTitleCount = count * 3;
      const text = JSON.stringify(title[j]);
      oWorksheet
        .GetRangeByNumber(nRowTitle + j, passedTitleCount)
        .SetValue(text);
      oWorksheet
        .GetRangeByNumber(nRowTitle + j, passedTitleCount)
        .AutoFit(false, true);
    }

    let nRowLink = 9;
    for (let h = 0; h < link.length; h++) {
      let passedLinkCount = count * 3 + 1;
      const text = JSON.stringify(link[h]);
      const textWithoutQuotes = text.replace(/"/g,'');
      oWorksheet.GetRangeByNumber(nRowLink + h, passedLinkCount).SetValue(textWithoutQuotes);
      oWorksheet
        .GetRangeByNumber(nRowLink + h, passedLinkCount)
        .AutoFit(false, true);
    }
  }

  function reloadCellValues() {
    let reload = setTimeout(function () {
      Api.asc_calculate(Asc.c_oAscCalculateType.All);
    }, 5000);
  }

  function getUrl(count) {
    const rn = Number(oWorksheet.GetRange("B2").GetValue());
    const pn = Number(oWorksheet.GetRange("B3").GetValue());
    const noCache = oWorksheet.GetRange("B4").GetValue();

    const url = `http://localhost:3000/advancedSearch?query=${query[count]}&rn=${rn}&pn=${pn}&no_cache=${noCache}`;

    return url;
  }

  for (let count = 0; count < query.length; count++) {
    // const url = `http://localhost:3000/test?query=${query[count]}`; <-- This is the test route which returns pre-set output from sampleData.js
    const url = getUrl(count);
    const request = new XMLHttpRequest();
    request.open("GET", url);
    request.onload = function () {
      const data = JSON.parse(request.response);
      const { organic_results: results } = data;
      const title = results.map((i) => i.title);
      const link = results.map((i) => i.link);
      populate(count, title, link);
    };
    request.send();
  }
  reloadCellValues();
})();
