const spreadsheet = new Spreadsheet(10, 10);

const controller = new Controller(spreadsheet);

const painter = new Painter();

painter.init(controller.getSpreadsheet());