const spreadsheet = new Spreadsheet(100, 100);

const controller = new Controller(spreadsheet);

const painter = new Painter();

painter.init(controller.getSpreadsheet());