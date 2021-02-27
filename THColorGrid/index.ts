import {IInputs, IOutputs} from "./generated/ManifestTypes";

import DataSetInterfaces = ComponentFramework.PropertyHelper.DataSetApi;
type DataSet = ComponentFramework.PropertyTypes.DataSet;

// Define const here
const RowRecordId: string = "rowRecId";

// Style name of Load More Button
const LoadMoreButton_Hidden_Style = "LoadMoreButton_Hidden_Style";

class UserExpression {
    field:string = "";
    colors:Array<string> = [];
    compareTo:string="";
    valid : boolean = false;
}

export class THColorGrid implements ComponentFramework.StandardControl<IInputs, IOutputs> {
// Cached context object for the latest updateView
    private contextObj: ComponentFramework.Context<IInputs>;

    // Div element created as part of this control's main container
    private mainContainer: HTMLDivElement;

    // Table element created as part of this control's table
    private dataTable: HTMLTableElement;

    // Button element created as part of this control
    private loadPageButton: HTMLButtonElement;


    /**
     * Empty constructor.
     */
    constructor() {

    }

    /**
     * Used to initialize the control instance. Controls can kick off remote server calls and other initialization actions here.
     * Data-set values are not initialized here, use updateView.
     * @param context The entire property bag available to control via Context Object; It contains values as set up by the customizer mapped to property names defined in the manifest, as well as utility functions.
     * @param notifyOutputChanged A callback method to alert the framework that the control has new outputs ready to be retrieved asynchronously.
     * @param state A piece of data that persists in one session for a single user. Can be set at any point in a controls life cycle by calling 'setControlState' in the Mode interface.
     * @param container If a control is marked control-type='starndard', it will receive an empty div element within which it can render its content.
     */
    public init(context: ComponentFramework.Context<IInputs>, notifyOutputChanged: () => void, state: ComponentFramework.Dictionary, container: HTMLDivElement) {
        // Need to track container resize so that control could get the available width. The available height won't be provided even this is true
        context.mode.trackContainerResize(true);

        // Create main table container div. 
        this.mainContainer = document.createElement("div");
        this.mainContainer.classList.add("SimpleTable_MainContainer_Style");

        // Create data table container div. 
        this.dataTable = document.createElement("table");
        this.dataTable.classList.add("SimpleTable_Table_Style");

        // Create data table container div. 
        this.loadPageButton = document.createElement("button");
        this.loadPageButton.setAttribute("type", "button");
        this.loadPageButton.innerText = context.resources.getString("PCF_ColorGrid_LoadMore_ButtonLabel");
        this.loadPageButton.classList.add(LoadMoreButton_Hidden_Style);
        this.loadPageButton.classList.add("LoadMoreButton_Style");
        this.loadPageButton.addEventListener("click", this.onLoadMoreButtonClick.bind(this));

        // Adding the main table and loadNextPage button created to the container DIV.
        this.mainContainer.appendChild(this.dataTable);
        this.mainContainer.appendChild(this.loadPageButton);
        container.appendChild(this.mainContainer);
    }


    /**
     * Called when any value in the property bag has changed. This includes field values, data-sets, global values such as container height and width, offline status, control metadata values such as label, visible, etc.
     * @param context The entire property bag available to control via Context Object; It contains values as set up by the customizer mapped to names defined in the manifest, as well as utility functions
     */
    public updateView(context: ComponentFramework.Context<IInputs>): void {
        this.contextObj = context;
        this.toggleLoadMoreButtonWhenNeeded(context.parameters.dataset);

        if (!context.parameters.dataset.loading) {

            // Get sorted columns on View
            let columnsOnView = this.getSortedColumnsOnView(context);

            if (!columnsOnView || columnsOnView.length === 0) {
                return;
            }

            let columnWidthDistribution = this.getColumnWidthDistribution(context, columnsOnView);


            while (this.dataTable.firstChild) {
                this.dataTable.removeChild(this.dataTable.firstChild);
            }

            this.dataTable.appendChild(this.createTableHeader(
                columnsOnView, 
                columnWidthDistribution, 
                context.parameters.headerBackgroundColor.raw, 
                context.parameters.headerForegroundColor.raw));

            this.dataTable.appendChild(this.createTableBody(this, columnsOnView, 
                columnWidthDistribution, 
                context.parameters.dataset, 
                context.parameters.cellBackgroundColor.raw, 
                context.parameters.rowBackColor.raw));


            this.dataTable.parentElement!.style.height = window.innerHeight - this.dataTable.offsetTop - 70 + "px";
        }
    }
    /** 
     * It is called by the framework prior to a control receiving new data. 
     * @returns an object based on nomenclature defined in manifest, expecting object[s] for property marked as “bound” or “output”
     */
    public getOutputs(): IOutputs {
        return {};
    }

    /** 
         * Called when the control is to be removed from the DOM tree. Controls should use this call for cleanup.
     * i.e. cancelling any pending remote calls, removing listeners, etc.
     */
    public destroy(): void {
    }

    /**
     * Get sorted columns on view
     * @param context 
     * @return sorted columns object on View
     */
    private getSortedColumnsOnView(context: ComponentFramework.Context<IInputs>): DataSetInterfaces.Column[] {
        if (!context.parameters.dataset.columns) {
            return [];
        }

        let columns = context.parameters.dataset.columns
            .filter(function (columnItem: DataSetInterfaces.Column) {
                // some column are supplementary and their order is not > 0
                return columnItem.order >= 0
            }
            );

        // Sort those columns so that they will be rendered in order
        columns.sort(function (a: DataSetInterfaces.Column, b: DataSetInterfaces.Column) {
            return a.order - b.order;
        });

        return columns;
    }

    /**
     * Get column width distribution
     * @param context context object of this cycle
     * @param columnsOnView columns array on the configured view
     * @returns column width distribution
     */
    private getColumnWidthDistribution(context: ComponentFramework.Context<IInputs>, columnsOnView: DataSetInterfaces.Column[]): string[] {

        let widthDistribution: string[] = [];

        // Considering need to remove border & padding length
        let totalWidth: number = context.mode.allocatedWidth - 250;
        let widthSum = 0;

        columnsOnView.forEach(function (columnItem) {
            widthSum += columnItem.visualSizeFactor;
        });

        let remainWidth: number = totalWidth;

        columnsOnView.forEach(function (item, index) {
            let widthPerCell = "";
            if (index !== columnsOnView.length - 1) {
                let cellWidth = Math.round((item.visualSizeFactor / widthSum) * totalWidth);
                remainWidth = remainWidth - cellWidth;
                widthPerCell = cellWidth + "px";
            }
            else {
                widthPerCell = remainWidth + "px";
            }
            widthDistribution.push(widthPerCell);
        });

        return widthDistribution;

    }


    private createTableHeader(columnsOnView: DataSetInterfaces.Column[], 
        widthDistribution: string[], 
        headerBackgroundColor: string, 
        headerForegroundColor: string): HTMLTableSectionElement {

        let tableHeader: HTMLTableSectionElement = document.createElement("thead");
        let tableHeaderRow: HTMLTableRowElement = document.createElement("tr");
        tableHeaderRow.classList.add("SimpleTable_TableRow_Style");

        tableHeaderRow.style.backgroundColor = headerBackgroundColor;
        tableHeaderRow.style.color = headerForegroundColor;

        columnsOnView.forEach(function (columnItem, index) {
            let tableHeaderCell = document.createElement("th");
            tableHeaderCell.classList.add("SimpleTable_TableHeader_Style");
            let innerDiv = document.createElement("div");
            innerDiv.classList.add("SimpleTable_TableCellInnerDiv_Style");
            innerDiv.style.maxWidth = widthDistribution[index];
            innerDiv.innerText = columnItem.displayName;
            tableHeaderCell.appendChild(innerDiv);
            tableHeaderRow.appendChild(tableHeaderCell);
        });

        tableHeader.appendChild(tableHeaderRow);
        return tableHeader;
    }

    private createTableBody(thisControl: any, columnsOnView: DataSetInterfaces.Column[], widthDistribution: string[], gridParam: DataSet, cellBackgroundColorExpression: string, rowBackgroundColor: string): HTMLTableSectionElement {

        // retrieve cellBackGround and JSON process
        // check if the passing fieldColorExpressionis JSON object in this format 
        // {"field":"pro_yeartous","colors":["red","white","yellow"], "compareTo":"today"}
        let json:UserExpression = this.convertToJSON(cellBackgroundColorExpression);
        let fieldValue:string = ""; //for storing value retrieve from field bound to the column

        let tableBody: HTMLTableSectionElement = document.createElement("tbody");
        let rowNumber: number = 0;

        if (gridParam.sortedRecordIds.length > 0) {
            for (let currentRecordId of gridParam.sortedRecordIds) {

                let tableRecordRow: HTMLTableRowElement = document.createElement("tr");
                tableRecordRow.classList.add("SimpleTable_TableRow_Style");
                let rowBackgroundColor1 = 'lightblue';

                try {
                    // rowBackgroundColor1 = "=((ROWNUMBER() % 2) == 0) ? 'lightblue' : 'white';"
                    if (rowBackgroundColor.startsWith("=")) {
                        let rowBackColorExpression = rowBackgroundColor.substring(1);
                        rowBackColorExpression = rowBackColorExpression.replace('ROWNUMBER()', 'rowNumber');
                        rowBackgroundColor1 = eval(rowBackColorExpression);
                    }
                    else {
                        rowBackgroundColor1 = rowBackgroundColor;
                    }
                    console.log(rowBackgroundColor1);

                    tableRecordRow.style.backgroundColor = rowBackgroundColor1;
                }
                catch (err) {
                    console.log("Invalid Row Background Color expression");
                    console.log(err.message);
                }
                rowNumber++;


                tableRecordRow.addEventListener("click", this.onRowClick.bind(this));

                // Set the recordId on the row dom
                tableRecordRow.setAttribute(RowRecordId, gridParam.records[currentRecordId].getRecordId());


                let thisClass = this; //this object will lost context when come inside forEach

                columnsOnView.forEach(function (columnItem, index) {
                    let tableRecordCell = document.createElement("td");
                    tableRecordCell.classList.add("SimpleTable_TableCell_Style");


                    //Single Color Expression for background color implementation. 
                    var fieldcolorExpression = cellBackgroundColorExpression;
                    var fieldname = '';
                    //var backgroundColor = 'white';
                    //Kien: by default cellbackground color should be the same as row background color
                    var backgroundColor = rowBackgroundColor1

                    //new code starts here                  
                    if (json.valid && columnItem.name == json.field) {
                        try {
                            let val = gridParam.records[currentRecordId].getValue(json.field);
                            fieldValue = val.toString();
                            console.log("field-value retrieve from CRM:", fieldValue);

                            let today = thisClass.getDateWithoutTime(new Date());
                            let fieldValueAsDate = thisClass.getDateWithoutTime(new Date(fieldValue));
                            
                            let todayAsInt = today.getTime();            //turn date to double-int
                            let fieldAsInt= fieldValueAsDate.getTime();    //turn date to double-in to compare

                            if (fieldAsInt == todayAsInt) {
                                console.log(`field:${fieldAsInt} = ${todayAsInt}`);  //for debug
                                backgroundColor = json.colors[1];
                            } else if (fieldAsInt < todayAsInt) {
                                console.log(`field:${fieldAsInt} < ${todayAsInt}`);  //for debug
                                backgroundColor = json.colors[0];
                            } else {
                                console.log(`field:${fieldAsInt} > ${todayAsInt}`);  //for debug
                                backgroundColor = json.colors[2];
                            }

                        } catch (e){
                            console.log(`error when reading value from field (${json.field}): ${e.message}`);
                            //if error,  use default background
                        }
                    }
                    else //new code end here



                    //check whether it is an expression. 
                    if (fieldcolorExpression.startsWith("=")) {
                        //Check whether the expression has fieldname
                        if (fieldcolorExpression.indexOf(fieldname) > -2) {
                            //Remove =
                            fieldcolorExpression = fieldcolorExpression.replace('=', '');

                            //Get field name
                            let start = fieldcolorExpression.indexOf("('") + 2;
                            let finish = fieldcolorExpression.indexOf("')");
                            let fieldname = fieldcolorExpression.substring(start, finish);
                            console.log(fieldname);

                            if (columnItem.name == fieldname) {
                                //Replace GETFIELDVALUE by dataset getValue
                                fieldcolorExpression = fieldcolorExpression.replace('GETFIELDVALUE', 'gridParam.records[currentRecordId].getValue');

                                console.log(fieldcolorExpression);

                                try {
                                    backgroundColor = eval(fieldcolorExpression);
                                    //console.log(backgroundColor);
                                    //tableRecordCell.style.backgroundColor = backgroundColor;

                                }
                                catch (err) {
                                    console.log("Invalid field color expression" + fieldcolorExpression);
                                    console.log(err.message);
                                    return;
                                }

                            }
                        }
                    }
                    else {//No expression,user has set hardcoded color in property
                        if (fieldcolorExpression != '') {
                            backgroundColor = fieldcolorExpression;
                            //console.log(backgroundColor);
                            //tableRecordCell.style.backgroundColor = backgroundColor;
                        }
                    }

                    //now we have background color, use it to style the cell
                    console.log(backgroundColor);
                    tableRecordCell.style.backgroundColor = backgroundColor;



                    let innerDiv = document.createElement("div");
                    innerDiv.classList.add("SimpleTable_TableCellInnerDiv_Style");
                    innerDiv.style.maxWidth = widthDistribution[index];
                    innerDiv.innerText = gridParam.records[currentRecordId].getFormattedValue(columnItem.name);
                    tableRecordCell.appendChild(innerDiv);
                    tableRecordRow.appendChild(tableRecordCell);
                });

                tableBody.appendChild(tableRecordRow);
            }
        }
        else {
            let tableRecordRow: HTMLTableRowElement = document.createElement("tr");
            let tableRecordCell: HTMLTableCellElement = document.createElement("td");
            tableRecordCell.classList.add("No_Record_Style");
            tableRecordCell.colSpan = columnsOnView.length;
            tableRecordCell.innerText = this.contextObj.resources.getString("PCF_ColorGrid_No_Record_Found");
            tableRecordRow.appendChild(tableRecordCell)
            tableBody.appendChild(tableRecordRow);
        }

        return tableBody;
    }

 
    /**
     * Row Click Event handler for the associated row when being clicked
     * @param event
     */
    private onRowClick(event: Event): void {
        let rowRecordId = (event.currentTarget as HTMLTableRowElement).getAttribute(RowRecordId);

        if (rowRecordId) {
            let entityReference = this.contextObj.parameters.dataset.records[rowRecordId].getNamedReference();
            let entityFormOptions = {
                entityName: entityReference.entityType!,
                entityId: entityReference.id,
            }
            this.contextObj.navigation.openForm(entityFormOptions);
        }
    }

    /**
     * Toggle 'LoadMore' button when needed
     */
    private toggleLoadMoreButtonWhenNeeded(gridParam: DataSet): void {

        if (gridParam.paging.hasNextPage && this.loadPageButton.classList.contains(LoadMoreButton_Hidden_Style)) {
            this.loadPageButton.classList.remove(LoadMoreButton_Hidden_Style);
        }
        else if (!gridParam.paging.hasNextPage && !this.loadPageButton.classList.contains(LoadMoreButton_Hidden_Style)) {
            this.loadPageButton.classList.add(LoadMoreButton_Hidden_Style);
        }

    }

    /**
     * 'LoadMore' Button Event handler when load more button clicks
     * @param event
     */
    private onLoadMoreButtonClick(event: Event): void {
        this.contextObj.parameters.dataset.paging.loadNextPage();
        this.toggleLoadMoreButtonWhenNeeded(this.contextObj.parameters.dataset);
    }


    /** convert the provide Date value into JavaScript date only object (no time) */
    private getDateWithoutTime(d:Date): Date {
        return new Date( d.getFullYear(), d.getMonth(), d.getDate());
    }

    /** 
     * convert a provided epxression to UserExpression type object.
     * 
     */
    private convertToJSON(expression:string):UserExpression {
        let valid = false;
        let json = new UserExpression();

        try {
            if (typeof expression != 'string' || expression==null || expression.length==0)
                throw new Error("expression is not string, null or empty"); 

            json = JSON.parse(expression);    //convert string to JS object;

            //validate the json object, 
            if ( !json.hasOwnProperty("field") || json.compareTo.length==0 )
                throw new Error("passing string must have attribute {field}");
            
            if ( !json.hasOwnProperty("colors") || json.colors.length <3 )
                throw new Error("passing string must have attribute colors or must be array of three colors");

            if ( !json.hasOwnProperty("compareTo") || json.compareTo.length==0)
                throw new Error("passing string must have attribute compareTo");

            if (!json.compareTo.match("today") )      //compare string
                throw new Error("Sorry!, only support compareTo:'today' ");
                //later you can allow more options for compareTo such as : thismonth, thisyear, contains(regex)
            
            //set flag that indicate this is a valid object
            json.valid=true;
        }
        catch (er){
            console.log("convert expression to json encouter error:"+er.message);
            json.valid = false;
        }

        return json;
    }

}