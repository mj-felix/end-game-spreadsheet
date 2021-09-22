/** Helper class holding utility functions. */
class Utils {
    /**
     * Array holding characters of the alphabet.
     * 
     * @static
     * @type {Array.<string>}
     */
    static chars = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z'];

    /**
    *  Converts numeric index to Excel-like column header.
    * 
    *  @static
    *  @param {number} index Numeric column index.
    *  @returns {string} Column header.
    */
    static numberToExcelHeader(index) {
        if (index === 0) return '0';

        index -= 1;

        const quotient = Math.floor(index / 26);
        if (quotient > 0) {
            return Utils.numberToExcelHeader(quotient) + Utils.chars[index % 26];
        }

        return Utils.chars[index % 26];
    };

    /**
    *  Calculates result of the operation.
    *
    *  @static
    *  @param {string} operation Operation string.
    *  @returns {number} Result of operation.
    */
    static calculate(operation) {
        return Function(`return (${operation})`)();
    }
}

/** Class representing a Cell of data. */
class Cell {
    /**
     * Creates a Cell.
     * @constructor
     * @param {string} id Cell id; concatenated identifier of a column and row, e.g. A2.
     * @param {string} type Cell type; 'label' denotes header cell, 'input' denotes cell holding non-label value.
     * @param {string} value Cell value; input entered into the cell for 'input' cells that do not start with =, label for 'label' cells.
     */
    constructor(id, type = 'input', value = null) {
        /**
         * Cell id; concatenated identifier of a column and row, e.g. A2.
         * @type {string}
         */
        this.id = id;
        /**
         * Cell type; 'label' denotes header cell, 'input' denotes cell holding non-label value.
         * @type {string}
         */
        this.type = type;
        /**
         * Cell value; input entered into the cell for 'input' cells that do not start with =, label for 'label' cells.
         * @type {string}
         */
        this.value = value;
        /**
         * Cell formula; input entered into the cell starting with =.
         * @type {string}
         */
        this.formula = null;
        /**
         * Ids of cells that use this cell in their formulas.
         * @type {Array.<string>}
         */
        this.impactedCellIds = [];
        /**
         * Ids of cells that this cell uses to calculate its value basing on its formula.
         * @type {Array.<string>}
         */
        this.reliantOnCellIds = [];
        /**
         * Bold flag; denotes if the value of the cell should be presented in bold in UI.
         * @type {boolean}
         */
        this.isBold = false;
        /**
         * Italic flag; denotes if the value of the cell should be presented in italic in UI.
         * @type {boolean}
         */
        this.isItalic = false;
        /**
         * Underlined flag; denotes if the value of the cell should be presented as underlined in UI.
         * @type {boolean}
         */
        this.isUnderlined = false;
    }

    /**
    *  Returns Cell id.
    *
    *  @returns {string} Cell id.
    */
    getId() {
        return this.id;
    }

    /**
    *  Returns Cell value.
    *
    *  @returns {string} Cell value.
    */
    getValue() {
        return this.value;
    }

    /**
    *  Sets Cell value.
    *
    *  @param {string} value Cell value.
    */
    setValue(value) {
        this.value = value;
    }

    /**
    *  Returns Cell folrmula.
    *
    *  @returns {string} Cell formula.
    */
    getFormula() {
        return this.formula;
    }

    /**
    *  Sets Cell formula.
    *
    *  @param {string} formula Cell formula.
    */
    setFormula(formula) {
        this.formula = formula && formula.toUpperCase();
    }

    /**
    *  Returns Bold flag.
    *
    *  @returns {boolean} Bold flag.
    */
    getIsBold() {
        return this.isBold;
    }

    /**
    *  Toggles Bold flag and returns new value of Bold flag.
    *
    *  @returns {boolean} Bold flag.
    */
    toggleIsBold() {
        this.isBold = !this.isBold;
        return this.isBold;
    }

    /**
    *  Returns Italic flag.
    *
    *  @returns {boolean} Italic flag.
    */
    getIsItalic() {
        return this.isItalic;
    }

    /**
    *  Toggles Italic flag and returns new value of Italic flag.
    *
    *  @returns {boolean} Italic flag.
    */
    toggleIsItalic() {
        this.isItalic = !this.isItalic;
        return this.isItalic;
    }

    /**
    *  Returns Underlined flag.
    *
    *  @returns {boolean} Underlined flag.
    */
    getIsUnderlined() {
        return this.isUnderlined;
    }

    /**
    *  Toggles Underlined flag and returns new value of Underlined flag.
    *
    *  @returns {boolean} Underlined flag.
    */
    toggleIsUnderlined() {
        this.isUnderlined = !this.isUnderlined;
        return this.isUnderlined;
    }

    /**
    *  Adds Cell id into Impacted Cell Ids array.
    *
    *  @param {string} otherCellId Cell id.
    */
    addImpactedCell(otherCellId) {
        this.impactedCellIds.push(otherCellId);
    }

    /**
    *  Calculates and stores value of the cell when formula inputted is a sum.
    */
    updateValueFromSumFormula() {
        let marginalCells = this.formula.slice(5);
        marginalCells = marginalCells.substring(0, marginalCells.length - 1);
        marginalCells = marginalCells.split(':');

        const cellsInRange = spreadsheet.getMarginalCellsForRange(marginalCells);


        let result = 0;

        try {

            for (const cellId of cellsInRange) {
                let cell = spreadsheet.getCellById(cellId);
                result += parseFloat(cell.getValue() || 0);
                if (!cell.impactedCellIds.filter((impactedCellId) => impactedCellId === this.getId()).length) {
                    cell.addImpactedCell(this.getId());
                }
            }

            this.value = Number.isNaN(result) ? '\'' + this.formula : result;
            this.reliantOnCellIds = cellsInRange;
        } catch (e) {
            console.log(e);
            this.value = '\'' + this.formula;
        }
    }

    /**
    *  Calculates and stores value of the cell when formula inputted consists of only addition, subtraction, mutiplication, division operators and round parentheses and cell ids.
    */
    updateValueFromRegularFormula() {
        let operation = this.formula.slice(1);
        const cellIds = operation.split(/[+\-*/()]/);

        const newCells = [];
        for (let cellId of cellIds) {
            let cell;
            try {
                cell = spreadsheet.getCellById(cellId);
            } catch (e) {
                continue;
            }
            const cellValue = cell.getValue();

            newCells.push([cellId, cellValue]);

            if (!cell.impactedCellIds.filter((impactedCellId) => impactedCellId === this.getId()).length) {
                cell.addImpactedCell(this.getId());
            }
        }

        // fix for A1=1, A10=1, B1=A1+A10 => 2 not 11
        newCells.sort((a, b) => b[0].length - a[0].length);
        // alternative solution with regex lookahead - replace next for loop with this:
        // for (let i = 0; i < newCells.length; i++) {
        //     const cellRegex = new RegExp(`${newCells[i][0]}(?![0-9])`, 'g');
        //     operation = operation.replace(cellRegex, newCells[i][1]);
        // }

        for (let i = 0; i < newCells.length; i++) {
            operation = operation.replaceAll(newCells[i][0], newCells[i][1]);
        }

        try {
            this.value = Utils.calculate(operation);
            this.reliantOnCellIds = newCells.map((newCell) => newCell[0]);
        } catch (e) {
            console.log(e);
            this.value = '\'' + this.formula;
        }
    }

    /**
    *  Calculates and stores value of the cell when formula inputted .
    */
    updateValueFromFormula() {
        if (this.formula.startsWith('=SUM')) {
            this.updateValueFromSumFormula();
        } else {
            this.updateValueFromRegularFormula();
        }
    }
}

/** Class representing a Spreadsheet of Cells (Model). */
class Spreadsheet {
    /**
     * Creates a Spreadsheet.
     * @constructor
     * @param {number} numOfColumns Number of columns for the spreadsheet.
     * @param {number} numOfRows Number of rows for the spreadsheet.
     */
    constructor(numOfColumns, numOfRows) {
        /**
         * Data object holding properties (indicating columns of spreadsheet) of arrays of Cells.
         * @type {object}
         */
        this.data = {};

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

    /**
    *  Returns Data object.
    *
    *  @returns {object} Data object.
    */
    getData() {
        return this.data;
    }

    /**
    *  Returns Cell object.
    *
    *  @param {string} column Spreadsheet column identifier.
    *  @param {string} row Spreadsheet row identifier.
    *  @returns {Cell} Cell object.
    */
    getCell(column, row) {
        return this.data[column][row];
    }

    /**
    *  Returns Cell object.
    *
    *  @param {string} columnRow Cell id.
    *  @returns {Cell} Cell object.
    */
    getCellById(columnRow) {
        const index = columnRow.search(/\d/);
        const column = columnRow.slice(0, index);
        const row = columnRow.slice(index);
        return this.getCell(column, row);
    }

    /**
    *  Returns spreadsheet column identifiers.
    *
    *  @returns {Array.<string>} Array of spreadsheet column identifiers.
    */
    getColumns() {
        return Object.keys(this.data);
    }

    /**
    *  Returns arrays of cell id strings (as restricted by provided 2 cell ids).
    *
    *  @param {Array.<string>} marginalCells Array of 2 cell ids.
    *  @returns {Array.<string>} Array of cell ids.
    */
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

        const cellIdsInRange = [];

        for (let i = firstColumnIndex; i <= lastColumnIndex; i++) {
            for (let j = firstRow; j <= lastRow; j++) {
                const newCellId = spreadsheetColumns[i] + j;
                cellIdsInRange.push(newCellId);
            }
        }

        return cellIdsInRange;

    }

}

/** Class representing a Painter (View). */
class Painter {
    /**
     * Creates a Painter.
     * @constructor
     */
    constructor() {
        /**
         * Active (focused on) cell id.
         * @type {string}
         */
        this.activeCellId = null;
    }

    /**
    *  Returns active cell id.
    *
    *  @returns {string} Cell id.
    */
    getActiveCellId() {
        return this.activeCellId;
    }

    /**
    *  Sets active cell id.
    *
    *  @param {string} cellId Cell id.
    */
    setActiveCellId(cellId) {
        this.activeCellId = cellId;
    }

    /**
    *  Initializes the spreadsheet: paints it and adds various listeners to listen to user actions.
    *
    *  @param {Spreadsheet} spreadsheet Spreadsheet of data.
    */
    init(spreadsheet) {
        this.paint(spreadsheet);
        this.addRefreshListener(spreadsheet);
        this.addBoldListener();
        this.addItalicListener();
        this.addUnderlineListener();
        this.addSetInactiveCellListener();
    }

    /**
    *  Adds listener to user clicking refresh button.
    *
    *  @param {Spreadsheet} spreadsheet Spreadsheet of data.
    */
    addRefreshListener(spreadsheet) {
        document.querySelector('#refreshButton').addEventListener('click', (event) => {
            document.getElementById('spreadsheet').innerHTML = 'Refreshing ...';
            setTimeout(() => this.paint(spreadsheet), 1000);
        });
    }

    /**
    *  Adds listener to user clicking bold button.
    */
    addBoldListener() {
        document.querySelector('#boldButton').addEventListener('click', (event) => {
            const activeCellId = this.getActiveCellId();
            if (activeCellId) {
                controller.toggleBold(activeCellId);
                document.getElementById(activeCellId).focus();
            }
        });
    }

    /**
    *  Adds listener to user clicking italic button.
    */
    addItalicListener() {
        document.querySelector('#italicButton').addEventListener('click', (event) => {
            const activeCellId = this.getActiveCellId();
            if (activeCellId) {
                controller.toggleItalic(activeCellId);
                document.getElementById(activeCellId).focus();
            }
        });
    }

    /**
    *  Adds listener to user clicking underline button.
    */
    addUnderlineListener() {
        document.querySelector('#underlineButton').addEventListener('click', (event) => {
            const activeCellId = this.getActiveCellId();
            if (activeCellId) {
                controller.toggleUnderlined(activeCellId);
                document.getElementById(activeCellId).focus();
            }
        });
    }

    /**
    *  Adds listener to user clicking outside of spreadsheet and not bold, italic, underline button to set active cell id to null.
    */
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

    /**
    *  Paints the spreadsheet.
    *
    *  @param {Spreadsheet} spreadsheet Spreadsheet of data.
    */
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

    /**
    *  Adds cells listeners.
    *
    *  @param {Spreadsheet} spreadsheet Spreadsheet of data.
    */
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
                controller.updateValueFromFormulaAndCellReferences(event.target.id, event.target.value);
            });
            cell.addEventListener('keyup', (event) => {
                if (event.keyCode === 13) {
                    this.moveCursorToCellBelow(event, spreadsheet);
                }
            });
        }

    }

    /**
    *  Moves cursor to cell below or skips to the top of next column.
    *
    *  @param {object} event Event triggered.
    *  @param {Spreadsheet} spreadsheet Spreadsheet of data.
    */
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

    /**
    *  Makes cell bold in UI.
    *
    *  @param {string} cellId Cell id.
    *  @param {boolean} isBold Bold flag.
    */
    setBold(cellId, isBold) {
        const inputClassList = document.getElementById(cellId).classList;
        if (isBold) {
            inputClassList.add('strong');
        } else {
            inputClassList.remove('strong');
        }
    }

    /**
    *  Makes cell italic in UI.
    *
    *  @param {string} cellId Cell id.
    *  @param {boolean} isItalic Italic flag.
    */
    setItalic(cellId, isItalic) {
        const inputClassList = document.getElementById(cellId).classList;
        if (isItalic) {
            inputClassList.add('italic');
        } else {
            inputClassList.remove('italic');
        }
    }

    /**
    *  Makes cell underlined in UI.
    *
    *  @param {string} cellId Cell id.
    *  @param {boolean} isUnderlined Underlined flag.
    */
    setUnderlined(cellId, isUnderlined) {
        const inputClassList = document.getElementById(cellId).classList;
        if (isUnderlined) {
            inputClassList.add('underlined');
        } else {
            inputClassList.remove('underlined');
        }
    }

    /**
    *  Updates cell in UI to present new value.
    *
    *  @param {string} cellId Cell id.
    *  @param {string} newInputValue New input value.
    */
    updateInput(cellId, newInputValue) {
        document.getElementById(cellId).value = newInputValue;
    }

}

/** Class representing a Controller. */
class Controller {
    /**
     * Creates a Controller.
     * @constructor
     * @param {Spreadsheet} spreadsheet Spreadsheet of data.
     */
    constructor(spreadsheet) {
        /**
         * Spreadsheet of data
         * @type {Spreadsheet}
         */
        this.spreadsheet = spreadsheet;
    }

    /**
    *  Returns Spreadsheet of data.
    *
    *  @returns {Spreadsheet} Spreadsheet of data.
    */
    getSpreadsheet() {
        return this.spreadsheet;
    }

    /**
    *  Toggles cell's bold flag and updates UI accordignly.
    *
    *  @param {string} cellId Cell id.
    */
    toggleBold(cellId) {
        const cell = this.spreadsheet.getCellById(cellId);
        const isBold = cell.toggleIsBold();
        painter.setBold(cellId, isBold);
    }

    /**
    *  Toggles cell's italic flag and updates UI accordignly.
    *
    *  @param {string} cellId Cell id.
    */
    toggleItalic(cellId) {
        const cell = this.spreadsheet.getCellById(cellId);
        const isItalic = cell.toggleIsItalic();
        painter.setItalic(cellId, isItalic);
    }

    /**
    *  Toggles cell's underlined flag and updates UI accordignly.
    *
    *  @param {string} cellId Cell id.
    */
    toggleUnderlined(cellId) {
        const cell = this.spreadsheet.getCellById(cellId);
        const isUnderlined = cell.toggleIsUnderlined();
        painter.setUnderlined(cellId, isUnderlined);
    }

    /**
    *  Saves the cell's new value or formula.
    *
    *  @param {string} cellId Cell id.
    *  @param {string} value Cell id.
    */
    saveCell(cellId, value, isBlurEvent) {
        const cell = this.spreadsheet.getCellById(cellId);

        // if formula in the cell
        if (value.startsWith('=')) {
            cell.setFormula(value);
        } else { // if no formula in the cell
            cell.setValue(value);
            cell.setFormula(null);
        }

    }

    /**
    *  If cell is defocused with formula entered, calculates the result basing on that formula; also updates other cells impacted by the change in this cell.
    *
    *  @param {string} cellId Cell id.
    *  @param {string} value Cell id.
    */
    updateValueFromFormulaAndCellReferences(cellId, value) {
        const cell = this.spreadsheet.getCellById(cellId);

        // update all the cells that impact this cell now that this cell doesn't hold formula any more or formula has changed
        for (let reliantOnCellId of cell.reliantOnCellIds) {
            const cell = spreadsheet.getCellById(reliantOnCellId);
            cell.impactedCellIds = cell.impactedCellIds.filter((impactedCellId) => impactedCellId !== cellId);
        }
        cell.reliantOnCellIds = [];

        // if formula in the cell
        if (value.startsWith('=')) {
            cell.updateValueFromFormula();
            painter.updateInput(cellId, cell.getValue());
        }

        // update other cells that are impacted by this cell's change
        this.updateImpactedCells(cell);

    }

    /**
    *  Updates cells impacted by the cell with cell id passed in.
    *
    *  @param {string} cellId Cell id.
    */
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

    /**
    *  Check if cell holds a formula and updates UI accordingly.
    *
    *  @param {string} cellId Cell id.
    */
    checkForFormula(cellId) {
        painter.setActiveCellId(cellId);

        const cell = this.spreadsheet.getCellById(cellId);
        const cellFormula = cell.getFormula();
        if (cellFormula) {
            painter.updateInput(cellId, cellFormula);
        }
    }

}