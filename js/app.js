class Utils {
    static chars = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z'];

    static numberToExcelHeader(index) {
        if (index === 0) return '0';

        index -= 1;

        const quotient = Math.floor(index / 26);
        if (quotient > 0) {
            return Utils.numberToExcelHeader(quotient) + Utils.chars[index % 26];
        }

        return Utils.chars[index % 26];
    };

    static calculate(operation) {
        return Function(`return (${operation})`)();
    }
}

class Cell {
    constructor(id, type = 'input', value = null, formula = null) {
        this.id = id;
        this.type = type;
        this.value = value;
        this.formula = formula;
        this.impactedCellIds = [];
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

    addImpactedCell(otherCellId) {
        this.impactedCellIds.push(otherCellId);
    }

    updateValueFromSumFormula() {
        let marginalCells = this.formula.slice(5);
        marginalCells = marginalCells.substring(0, marginalCells.length - 1);
        marginalCells = marginalCells.split(':');

        const { firstRow, lastRow, firstColumnIndex, lastColumnIndex, spreadsheetColumns } = spreadsheet.getMarginalCellsForRange(marginalCells);

        const newCellIds = [];
        let result = 0;

        try {

            for (let i = firstColumnIndex; i <= lastColumnIndex; i++) {
                for (let j = firstRow; j <= lastRow; j++) {
                    const newCellId = spreadsheetColumns[i] + j;
                    let cell = spreadsheet.getCellById(newCellId);
                    newCellIds.push(newCellId);
                    result += parseFloat(cell.getValue() || 0);
                    if (!cell.impactedCellIds.filter((impactedCellId) => impactedCellId === this.getId()).length) {
                        cell.addImpactedCell(this.getId());
                    }
                }
            }

            this.value = Number.isNaN(result) ? '\'' + this.formula : result;
            this.reliantOnCellIds = newCellIds;
        } catch (e) {
            console.log(e);
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
            if (!cell.impactedCellIds.filter((impactedCellId) => impactedCellId === this.getId()).length) {
                cell.addImpactedCell(this.getId());
            }
        }

        for (let i = 0; i < newCellIds.length; i++) {
            operation = operation.replaceAll(newCellIds[i], values[i]);
        }

        try {
            this.value = Utils.calculate(operation);
            this.reliantOnCellIds = newCellIds;
        } catch (e) {
            console.log(e);
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
            const cellId = Utils.numberToExcelHeader(i);
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

    getMarginalCellsForRange(marginalCells) {

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

        const spreadsheetColumns = this.getColumns();

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

        return { firstRow, lastRow, firstColumnIndex, lastColumnIndex, spreadsheetColumns };

    }

}

class Painter {
    constructor() {
        this.activeCellId = null;
    }

    getActiveCellId() {
        return this.activeCellId;
    }

    setActiveCellId(cellId) {
        this.activeCellId = cellId;
    }

    init(spreadsheet) {
        this.paint(spreadsheet);
        this.addRefreshListener(spreadsheet);
        this.addBoldListener();
        this.addItalicListener();
        this.addUnderlineListener();
        this.addSetInactiveCellListener();
    }

    addRefreshListener(spreadsheet) {
        document.querySelector('#refreshButton').addEventListener('click', (event) => {
            document.getElementById('spreadsheet').innerHTML = 'Refreshing ...';
            setTimeout(() => this.paint(spreadsheet), 1000);
        });
    }

    addBoldListener() {
        document.querySelector('#boldButton').addEventListener('click', (event) => {
            const activeCellId = this.getActiveCellId();
            if (activeCellId) {
                controller.toggleBold(activeCellId);
                document.getElementById(activeCellId).focus();
            }
        });
    }

    addItalicListener() {
        document.querySelector('#italicButton').addEventListener('click', (event) => {
            const activeCellId = this.getActiveCellId();
            if (activeCellId) {
                controller.toggleItalic(activeCellId);
                document.getElementById(activeCellId).focus();
            }
        });
    }

    addUnderlineListener() {
        document.querySelector('#underlineButton').addEventListener('click', (event) => {
            const activeCellId = this.getActiveCellId();
            if (activeCellId) {
                controller.toggleUnderlined(activeCellId);
                document.getElementById(activeCellId).focus();
            }
        });
    }

    addSetInactiveCellListener() {
        document.body.addEventListener('click', (event) => {
            if (!event.target.classList.contains('cell')
                && event.target !== document.getElementById('boldButton')
                && event.target !== document.getElementById('italicButton')
                && event.target !== document.getElementById('underlineButton')
            ) {
                this.setActiveCellId(null);
            }
        });
    }

    paint(spreadsheet) {

        const data = spreadsheet.getData();
        const keys = spreadsheet.getColumns();
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

        this.listen(spreadsheet);
    }

    listen(spreadsheet) {

        const cells = document.querySelectorAll('input.cell');
        for (let cell of cells) {
            cell.addEventListener('input', (event) => {
                controller.saveCell(event.target.id, event.target.value);
            });
            cell.addEventListener('focus', (event) => {
                controller.checkForFormula(event.target.id, event);
            });
            cell.addEventListener('blur', (event) => {
                controller.saveCell(event.target.id, event.target.value, true);
            });
            cell.addEventListener('keyup', (event) => {
                if (event.keyCode === 13) {
                    this.moveCursorToCellBelow(event, spreadsheet);
                }
            });
        }

    }

    moveCursorToCellBelow(event, spreadsheet) {

        const cellId = event.target.id;
        const index = cellId.search(/\d/);
        let column = cellId.slice(0, index);
        let row = cellId.slice(index);

        const numOfRows = spreadsheet.getData()['0'].length;
        if (row >= numOfRows - 1) {
            row = 1;
            const keys = spreadsheet.getColumns();
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

    getSpreadsheet() {
        return this.spreadsheet;
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

        // if formula in the cell
        if (value.startsWith('=')) {
            cell.setFormula(value);
            // if exited the cell
            if (isBlurEvent) {
                cell.updateValueFromFormula();
                painter.updateInput(cellId, cell.getValue() || '\'' + cell.getFormula());
            }
        } else { // if no formula in the cell
            cell.setValue(value);
            cell.setFormula(null);
            // update all the cells that impact this cell now that this cell doesn't hold formula any more
            for (let reliantOnCellId of cell.reliantOnCellIds) {
                const cell = spreadsheet.getCellById(reliantOnCellId);
                cell.impactedCellIds = cell.impactedCellIds.filter((impactedCellId) => impactedCellId !== cellId);
            }
            cell.reliantOnCellIds = [];
        }

        // update other cells that are impacted by this cell's change
        if (isBlurEvent) {
            this.updateImpactedCells(cell);
        }

    }

    updateImpactedCells(cell) {
        for (let impactedCellId of cell.impactedCellIds) {
            // removed after introduction of cell.impactedCellIds
            // if (!impactedCell.formula) {
            //     // cell.impactedCells = cell.impactedCells.filter((cell) => cell.getId() !== impactedCell.getId());
            //     continue;
            // }
            const impactedCell = spreadsheet.getCellById(impactedCellId);
            impactedCell.updateValueFromFormula();
            painter.updateInput(impactedCell.getId(), impactedCell.getValue());
            if (impactedCell.impactedCellIds && impactedCell.impactedCellIds.length) {
                this.updateImpactedCells(impactedCell);
            }
        }
    }

    checkForFormula(cellId) {
        painter.setActiveCellId(cellId);

        const cell = this.spreadsheet.getCellById(cellId);
        const cellFormula = cell.getFormula();
        if (cellFormula) {
            painter.updateInput(cellId, cellFormula);
        }
    }

}