(function () {
  const oWorksheet = Api.GetActiveSheet();

  const putValues = () => {
    oWorksheet.GetRange("A1").SetValue("Queries");
    oWorksheet.GetRange("A1").SetBold(true);
    oWorksheet.GetRange("A2").SetValue("que1");
    oWorksheet.GetRange("A3").SetValue("que2");
    oWorksheet.GetRange("A4").SetValue("que3");
    for (let i = 1; i <= 4; i++) {
      oWorksheet.GetRange(`A${i}`).AutoFit(false, true);
      oWorksheet.GetRange(`A${i}`).SetAlignHorizontal("center");
    }

    oWorksheet.GetRange("B1").SetValue("Params");
    oWorksheet.GetRange("B1").SetBold(true);

    oWorksheet.GetRange("B2").SetValue("No. of Results");
    oWorksheet.GetRange("B3").SetValue("No. of Pages");
    oWorksheet.GetRange("B4").SetValue("No Cache? Y or N)");
    for (let i = 1; i <= 4; i++) {
      oWorksheet.GetRange(`B${i}`).AutoFit(false, true);
      oWorksheet.GetRange(`B${i}`).SetAlignHorizontal("center");
    }
  };

  const button = () => {
    var oWorksheet = Api.GetActiveSheet();
    var oFill = Api.CreateSolidFill(Api.CreateRGBColor(36, 160, 237));
    var oStroke = Api.CreateStroke(0, Api.CreateNoFill());
    var oShape = oWorksheet.AddShape(
      "rect",
      34 * 36000,
      15 * 36000,
      oFill,
      oStroke,
      3,
      0,
      1,
      0
    );
    var oDocContent = oShape.GetContent();
    var oParagraph = oDocContent.GetElement(0);
    var oParaPr = oParagraph.GetParaPr();
    oParaPr.SetJc("center");
    oParagraph.AddText("Assign the MAIN MACRO to this shape");
  };

  const formatData = () => {
    oWorksheet.GetRange("A8:B8").Merge(true);
    oWorksheet.GetRange("A8:B8").SetValue("Query1");
    oWorksheet.GetRange("A8").SetFontSize("18");

    oWorksheet.GetRange("A9").SetValue("Title");
    oWorksheet.GetRange("B9").SetValue("Link");
    oWorksheet.GetRange("A9").SetFontSize("14");
    oWorksheet.GetRange("B9").SetFontSize("14");

    oWorksheet.GetRange("A8:B8").SetBold(true);
    oWorksheet.GetRange("A8").SetAlignHorizontal("center");
    oWorksheet.GetRange("A9").SetAlignHorizontal("center");
    oWorksheet.GetRange("B9").SetAlignHorizontal("center");

    oWorksheet.GetRange("D8:E8").Merge(true);
    oWorksheet.GetRange("D8:E8").SetValue("Query2");
    oWorksheet.GetRange("D8").SetFontSize("18");

    oWorksheet.GetRange("D9").SetValue("Title");
    oWorksheet.GetRange("E9").SetValue("Link");
    oWorksheet.GetRange("D9").SetFontSize("14");
    oWorksheet.GetRange("E9").SetFontSize("14");

    oWorksheet.GetRange("D8:E8").SetBold(true);
    oWorksheet.GetRange("D8").SetAlignHorizontal("center");
    oWorksheet.GetRange("D9").SetAlignHorizontal("center");
    oWorksheet.GetRange("E9").SetAlignHorizontal("center");

    oWorksheet.GetRange("G8:H8").Merge(true);
    oWorksheet.GetRange("G8:H8").SetValue("Query3");
    oWorksheet.GetRange("G8").SetFontSize("18");

    oWorksheet.GetRange("G9").SetValue("Title");
    oWorksheet.GetRange("H9").SetValue("Link");

    oWorksheet.GetRange("G9").SetFontSize("14");
    oWorksheet.GetRange("H9").SetFontSize("14");
    oWorksheet.GetRange("G8:H8").SetBold(true);
    oWorksheet.GetRange("G8").SetAlignHorizontal("center");
    oWorksheet.GetRange("G9").SetAlignHorizontal("center");
    oWorksheet.GetRange("H9").SetAlignHorizontal("center");
  };

  putValues();
  formatData();
  button();
})();
