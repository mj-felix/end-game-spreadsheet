const chars = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z'];

const numberToExcelHeader = (index) => {
    if (index === 0) return '0';

    index -= 1;

    const quotient = Math.floor(index / 26);
    if (quotient > 0) {
        return numberToExcelHeader(quotient) + chars[index % 26];
    }

    return chars[index % 26];
};

class Cell {
    constructor(id, type = 'input', value = null, formula = null) {
        this.id = id;
        this.type = type;
        this.value = value;
        this.formula = formula;
    }

    getId() {
        return this.id;
    }

    getValue() {
        return this.value;
    }

    setValue(value) {
        this.value = value;
    }

    setFormula(formula) {
        this.formula = formula;
    }

}

class Spreadsheet {
    constructor(numOfColumns, numOfRows) {
        this.data = [];

        const firstColumn = [];
        firstColumn.push(new Cell('0', 'label', ''));
        for (let i = 1; i <= numOfRows; i++) {
            firstColumn.push(new Cell(i.toString(), 'label', i.toString()));
        }
        this.data['0'] = firstColumn;

        for (let i = 1; i <= numOfColumns; i++) {
            const cellId = numberToExcelHeader(i);
            const column = [];
            column.push(new Cell(cellId + 0, 'label', cellId));
            for (let j = 1; j <= numOfRows; j++) {
                column.push(new Cell(cellId + j));
            }
            this.data[cellId] = column;
        }

    }

    getData() {
        return this.data;
    }

    getCell(column, row) {
        return this.data[column][row];
    }

    getCellById(columnRow) {
        const index = columnRow.search(/\d/);
        const column = columnRow.slice(0, index);
        const row = columnRow.slice(index);
        return this.getCell(column, row);
    }

}

class Painter {
    constructor(spreadsheet) {
        this.spreadsheet = spreadsheet;
    }

    paint(targetElementId) {
        document.getElementById(targetElementId).innerHTML = '...';
        console.log('painter paints');
        const data = this.spreadsheet.getData();
        const keys = Object.keys(data);
        const numOfColumns = keys.length;
        const numOfRows = data['0'].length;
        const html = [];
        html.push('<table>');
        for (let i = 0; i < numOfRows; i++) {
            html.push('<tr>');
            for (let j = 0; j < numOfColumns; j++) {
                const cell = spreadsheet.getCell(keys[j], i);
                const cellId = cell.getId();
                const cellValue = cell.getValue() || '';
                const isTableHeader = (i === 0) || (j === 0);
                if (!isTableHeader) {
                    html.push(`<td><input id='${cellId}' type='text' value='${cellValue}'></td>`);
                }
                else {
                    html.push(`<th>${cellValue}</th>`);
                }
            }
            html.push('</tr>');
        }
        html.push('</table>');

        document.getElementById(targetElementId).innerHTML = html.join('');

        this.listen();
    }

    listen() {
        this.listenToInputEventsAndSaveCells();
    }

    listenToInputEventsAndSaveCells() {
        const cells = document.querySelectorAll('input');
        for (let cell of cells) {
            cell.addEventListener('input', (event) => {
                console.log(event.target.id, event.target.value);
                controller.saveCell(event.target.id, event.target.value);
            });
        }
    }

}

class Controller {
    constructor(spreadsheet) {
        this.spreadsheet = spreadsheet;
    }

    saveCell(cellId, value) {
        const cell = this.spreadsheet.getCellById(cellId);
        cell.setValue(value);
    }

}