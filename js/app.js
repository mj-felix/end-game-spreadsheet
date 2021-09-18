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

function calculate(operation) {
    return Function(`'use strict'; return (${operation})`)();
}

class Cell {
    constructor(id, type = 'input', value = null, formula = null) {
        this.id = id;
        this.type = type;
        this.value = value;
        this.formula = formula;
        this.impactedCells = [];
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

    getFormula() {
        return this.formula;
    }

    setFormula(formula) {
        this.formula = formula && formula.toUpperCase();
    }

    addImpactedCell(otherCell) {
        this.impactedCells.push(otherCell);
    }

    updateValueFromFormula() {
        console.log('updateValueFromFormula');
        let operation = this.formula.slice(1);
        const cellIds = operation.split(/[+\-*/()]/);
        console.log(cellIds);
        const newCellIds = [];
        const values = [];
        for (let cellId of cellIds) {
            let cell;
            try {
                cell = spreadsheet.getCellById(cellId);
            } catch (e) {
                continue;
            }
            newCellIds.push(cellId);
            const cellValue = cell.getValue();
            values.push(cellValue);
            if (!cell.impactedCells.filter((impactedCell) => impactedCell.getId() === this.getId()).length) {
                cell.addImpactedCell(this);
            }
        }

        for (let i = 0; i < newCellIds.length; i++) {
            operation = operation.replaceAll(newCellIds[i], values[i]);
        }
        console.log(operation);
        try {
            this.value = calculate(operation);
        } catch (e) {
            this.value = '\'' + this.formula;
        }
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

    init() {
        this.paint();
        this.listenToRefreshClickEventAndRepaint();
    }

    listenToRefreshClickEventAndRepaint() {
        document.querySelector('#refreshButton').addEventListener('click', (event) => {
            this.paint();
        });
    }

    paint() {
        console.log('Painter paints ...');
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
                    html.push(`<td><input id='${cellId}' class='cell' type='text' value='${cellValue}'></td>`);
                }
                else {
                    html.push(`<th>${cellValue}</th>`);
                }
            }
            html.push('</tr>');
        }
        html.push('</table>');

        document.getElementById('spreadsheet').innerHTML = html.join('');

        this.listen();
    }

    listen() {

        const cells = document.querySelectorAll('input.cell');
        for (let cell of cells) {
            cell.addEventListener('input', (event) => {
                console.log(event.target.id, event.target.value);
                controller.saveCell(event.target.id, event.target.value);
            });
            cell.addEventListener('focus', (event) => {
                controller.checkForFormula(event.target.id);
            });
            cell.addEventListener('blur', (event) => {
                controller.saveCell(event.target.id, event.target.value, true);
            });
            cell.addEventListener('keyup', (event) => {
                painter.moveCursorToCellBelow(event);
            });
        }
    }

    moveCursorToCellBelow(event) {
        if (event.keyCode === 13) {
            const cellId = event.target.id;
            const index = cellId.search(/\d/);
            let column = cellId.slice(0, index);
            let row = cellId.slice(index);

            const numOfRows = this.spreadsheet.getData()['0'].length;
            console.log(column, row, numOfRows - 1);
            if (row >= numOfRows - 1) {
                row = 1;
                const keys = Object.keys(this.spreadsheet.getData());
                let nextKey;
                for (let i = 1; i < keys.length; i++) {
                    if (keys[i] === column) {
                        nextKey = keys[i + 1];
                        break;
                    }
                }
                column = nextKey;
            } else {
                row = parseInt(row, 10) + 1;
            }
            try {
                document.getElementById(column + row).focus();
            } catch (e) { }
        }
    }

    updateInput(cellId, newInputValue) {
        document.getElementById(cellId).value = newInputValue;
    }

}

class Controller {
    constructor(spreadsheet) {
        this.spreadsheet = spreadsheet;
    }

    saveCell(cellId, value, isBlurEvent) {
        const cell = this.spreadsheet.getCellById(cellId);
        if (value.startsWith('=')) {
            cell.setFormula(value);
            if (isBlurEvent) {
                cell.updateValueFromFormula();
                painter.updateInput(cellId, cell.getValue());
                // this.updateImpactedCells(cell);
            }
        } else {
            cell.setValue(value);
            cell.setFormula(null);
        }
        if (isBlurEvent) {
            this.updateImpactedCells(cell);
        }

    }

    updateImpactedCells(cell) {
        for (let impactedCell of cell.impactedCells) {
            if (!impactedCell.formula) {
                cell.impactedCells = cell.impactedCells.filter((cell) => cell.getId() !== impactedCell.getId());
                continue;
            }
            console.log(impactedCell);
            impactedCell.updateValueFromFormula();
            painter.updateInput(impactedCell.getId(), impactedCell.getValue());
            if (impactedCell.impactedCells && impactedCell.impactedCells.length) {
                this.updateImpactedCells(impactedCell);
            }
        }
    }

    checkForFormula(cellId) {
        const cell = this.spreadsheet.getCellById(cellId);
        const cellFormula = cell.getFormula();
        if (cellFormula) {
            painter.updateInput(cellId, cellFormula);
        }
    }


}