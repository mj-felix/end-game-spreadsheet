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
        this.reliantOnCellIds = [];
        this.isBold = false;
        this.isItalic = false;
        this.isUnderlined = false;
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

    getIsBold() {
        return this.isBold;
    }

    toggleIsBold() {
        this.isBold = !this.isBold;
        return this.isBold;
    }

    getIsItalic() {
        return this.isItalic;
    }

    toggleIsItalic() {
        this.isItalic = !this.isItalic;
        return this.isItalic;
    }

    getIsUnderlined() {
        return this.isUnderlined;
    }

    toggleIsUnderlined() {
        this.isUnderlined = !this.isUnderlined;
        return this.isUnderlined;
    }

    addImpactedCell(otherCell) {
        this.impactedCells.push(otherCell);
    }

    updateValueFromSumFormula() {
        let marginalCells = this.formula.slice(5);
        marginalCells = marginalCells.substring(0, marginalCells.length - 1);
        marginalCells = marginalCells.split(':');
        console.log(marginalCells);
        const index1 = marginalCells[0].search(/\d/);
        const column1 = marginalCells[0].slice(0, index1);
        const row1 = marginalCells[0].slice(index1);
        const index2 = marginalCells[1].search(/\d/);
        const column2 = marginalCells[1].slice(0, index2);
        const row2 = marginalCells[1].slice(index2);
        let firstRow, lastRow;
        if (parseInt(row1) < parseInt(row2)) {
            firstRow = row1;
            lastRow = row2;
        } else {
            firstRow = row2;
            lastRow = row1;
        }
        console.log(firstRow, lastRow);
        const spreadsheetColumns = spreadsheet.getColumns();

        const column1Index = spreadsheetColumns.findIndex((column) => column === column1);
        const column2Index = spreadsheetColumns.findIndex((column) => column === column2);
        let firstColumnIndex, lastColumnIndex;
        if (column1Index < column2Index) {
            firstColumnIndex = column1Index;
            lastColumnIndex = column2Index;
        } else {
            firstColumnIndex = column2Index;
            lastColumnIndex = column1Index;
        }
        console.log(firstColumnIndex, lastColumnIndex);

        const newCellIds = [];
        let result = 0;

        try {

            for (let i = firstColumnIndex; i <= lastColumnIndex; i++) {
                for (let j = firstRow; j <= lastRow; j++) {
                    const newCellId = spreadsheetColumns[i] + j;
                    let cell = spreadsheet.getCellById(newCellId);

                    newCellIds.push(newCellId);
                    result += parseFloat(cell.getValue() || 0);
                    if (!cell.impactedCells.filter((impactedCell) => impactedCell.getId() === this.getId()).length) {
                        cell.addImpactedCell(this);
                    }

                }
            }

            this.value = Number.isNaN(result) ? '\'' + this.formula : result;
            this.reliantOnCellIds = newCellIds;
        } catch (e) {
            this.value = '\'' + this.formula;
        }
    }

    updateValueFromRegularFormula() {
        let operation = this.formula.slice(1);
        const cellIds = operation.split(/[+\-*/()]/);

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

        try {
            this.value = calculate(operation);
            this.reliantOnCellIds = newCellIds;
        } catch (e) {
            this.value = '\'' + this.formula;
        }
    }

    updateValueFromFormula() {
        if (this.formula.startsWith('=SUM')) {
            this.updateValueFromSumFormula();
        } else {
            this.updateValueFromRegularFormula();
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

    getColumns() {
        return Object.keys(this.data);
    }

}

class Painter {
    constructor(spreadsheet) {
        this.spreadsheet = spreadsheet;
        this.activeCellId = null;
    }

    getActiveCellId() {
        return this.activeCellId;
    }

    setActiveCellId(cellId) {
        this.activeCellId = cellId;
    }

    init() {
        this.paint();
        this.listenToRefreshClickEventAndRepaint();
        this.listenToBoldClickEventAndUpdate();
        this.listenToItalicClickEventAndUpdate();
        this.listenToUnderlineClickEventAndUpdate();
    }

    listenToRefreshClickEventAndRepaint() {
        document.querySelector('#refreshButton').addEventListener('click', (event) => {
            document.getElementById('spreadsheet').innerHTML = 'Refreshing ...';
            setTimeout(() => this.paint(), 1000);
        });
    }

    listenToBoldClickEventAndUpdate() {
        document.querySelector('#boldButton').addEventListener('click', (event) => {
            const activeCellId = this.getActiveCellId();
            if (activeCellId) {
                controller.toggleBold(activeCellId);
                document.getElementById(activeCellId).focus();
            }
        });
    }

    listenToItalicClickEventAndUpdate() {
        document.querySelector('#italicsButton').addEventListener('click', (event) => {
            const activeCellId = this.getActiveCellId();
            if (activeCellId) {
                controller.toggleItalic(activeCellId);
                document.getElementById(activeCellId).focus();
            }
        });
    }

    listenToUnderlineClickEventAndUpdate() {
        document.querySelector('#underlineButton').addEventListener('click', (event) => {
            const activeCellId = this.getActiveCellId();
            if (activeCellId) {
                controller.toggleUnderlined(activeCellId);
                document.getElementById(activeCellId).focus();
            }
        });
    }

    paint() {

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
                    html.push(`<td><input id='${cellId}' class='cell${cell.getIsBold() ? ' strong' : ''}${cell.getIsItalic() ? ' italic' : ''}${cell.getIsUnderlined() ? ' underlined' : ''}' type='text' value='${cellValue}'></td>`);
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
                controller.saveCell(event.target.id, event.target.value);
            });
            cell.addEventListener('focus', (event) => {
                controller.checkForFormula(event.target.id);
                painter.setActiveCellId(event.target.id);
            });
            cell.addEventListener('blur', (event) => {
                controller.saveCell(event.target.id, event.target.value, true);
            });
            cell.addEventListener('keyup', (event) => {
                if (event.keyCode === 13) {
                    painter.moveCursorToCellBelow(event);
                }
            });
        }

    }

    moveCursorToCellBelow(event) {

        const cellId = event.target.id;
        const index = cellId.search(/\d/);
        let column = cellId.slice(0, index);
        let row = cellId.slice(index);

        const numOfRows = this.spreadsheet.getData()['0'].length;
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

    setBold(cellId, isBold) {
        const inputClassList = document.getElementById(cellId).classList;
        if (isBold) {
            inputClassList.add('strong');
        } else {
            inputClassList.remove('strong');
        }
    }

    setItalic(cellId, isItalic) {
        const inputClassList = document.getElementById(cellId).classList;
        if (isItalic) {
            inputClassList.add('italic');
        } else {
            inputClassList.remove('italic');
        }
    }

    setUnderlined(cellId, isUnderlined) {
        const inputClassList = document.getElementById(cellId).classList;
        if (isUnderlined) {
            inputClassList.add('underlined');
        } else {
            inputClassList.remove('underlined');
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

    toggleBold(cellId) {
        const cell = this.spreadsheet.getCellById(cellId);
        const isBold = cell.toggleIsBold();
        painter.setBold(cellId, isBold);
    }

    toggleItalic(cellId) {
        const cell = this.spreadsheet.getCellById(cellId);
        const isItalic = cell.toggleIsItalic();
        painter.setItalic(cellId, isItalic);
    }

    toggleUnderlined(cellId) {
        const cell = this.spreadsheet.getCellById(cellId);
        const isUnderlined = cell.toggleIsUnderlined();
        painter.setUnderlined(cellId, isUnderlined);
    }

    saveCell(cellId, value, isBlurEvent) {
        const cell = this.spreadsheet.getCellById(cellId);
        if (value.startsWith('=')) {
            cell.setFormula(value);
            if (isBlurEvent) {
                cell.updateValueFromFormula();
                painter.updateInput(cellId, cell.getValue() || '\'' + cell.getFormula());

            }
        } else {
            cell.setValue(value);
            cell.setFormula(null);
            for (let reliantOnCellId of cell.reliantOnCellIds) {
                const cell = spreadsheet.getCellById(reliantOnCellId);
                cell.impactedCells = cell.impactedCells.filter((cell) => cell.getId() !== cellId);
            }
            cell.reliantOnCellIds = [];
        }

        // update other cells that relies on this cell
        if (isBlurEvent) {
            this.updateImpactedCells(cell);
        }

    }

    updateImpactedCells(cell) {
        for (let impactedCell of cell.impactedCells) {
            // if (!impactedCell.formula) {
            //     // cell.impactedCells = cell.impactedCells.filter((cell) => cell.getId() !== impactedCell.getId());
            //     continue;
            // }
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