import { IInputs, IOutputs } from "./generated/ManifestTypes";
import { cellRendererOverrides } from "./customizers/CellRendererOverrides";
import { cellEditorOverrides } from "./customizers/CellEditorOverrides";
import { PAOneGridCustomizer } from "./types";
import * as React from "react";
import { DraggableRowsGridRenderer } from "./customizers/GridRenderer";
import { DraggableCell } from "./Controls/DraggableCell";


type Record = {
    diana_sortableid: string;
    new_task_sortorder: number;
    // other properties of the entity
 }
export class PAGridCustomizer implements ComponentFramework.ReactControl<IInputs, IOutputs> {
    private theComponent: ComponentFramework.ReactControl<IInputs, IOutputs>;
    private notifyOutputChanged: () => void;    
    /**
     * Empty constructor.
     */
    constructor() { }
    

    /**
     * Used to initialize the control instance. Controls can kick off remote server calls and other initialization actions here.
     * Data-set values are not initialized here, use updateView.
     * @param context The entire property bag available to control via Context Object; It contains values as set up by the customizer mapped to property names defined in the manifest, as well as utility functions.
     * @param notifyOutputChanged A callback method to alert the framework that the control has new outputs ready to be retrieved asynchronously.
     * @param state A piece of data that persists in one session for a single user. Can be set at any point in a controls life cycle by calling 'setControlState' in the Mode interface.
     */
    
    public init(
        context: ComponentFramework.Context<IInputs>,
        notifyOutputChanged: () => void,
        state: ComponentFramework.Dictionary

    ): void {
        this.notifyOutputChanged = notifyOutputChanged;
        const eventName = context.parameters.EventName.raw;

        function SortOrderChanged(executionContext: Xrm.Events.EventContext) {
            
            console.log("SortOrderChanged");
            const retrieveRecords = (sourceIndex: number, targetIndex: number) => {
               const parentId = executionContext.getFormContext().data.entity.getId();
               return Xrm.WebApi.retrieveMultipleRecords("diana_sortable", 
                  `?$select=diana_sortableid,new_task_sortorder&$filter=(_diana_accountid_value eq '${parentId}'` +
                  `and Microsoft.Dynamics.CRM.Between(PropertyName='new_task_sortorder',PropertyValues=['${sourceIndex}','${targetIndex}']))&$orderby=new_task_sortorder asc`)
            }
         
            const refreshGrid = () => {
                const formContext = executionContext.getFormContext();
                console.log(formContext.ui); // Log the ui property to the console
                console.log(formContext.ui.controls); // Log the controls property to the console
              
                //formContext.ui.controls.get("Sortables")?.refresh();
                Xrm.Utility.closeProgressIndicator();
            }
        
             
            const move = (sourceId: string, sourceValue: number, targetId: any, targetValue: number)=> {
                Xrm.Utility.showProgressIndicator("updating sort order");
             
                retrieveRecords(Math.min(sourceValue, targetValue), Math.max(sourceValue, targetValue)).then((response)=>{
                   const delta = sourceValue<targetValue ? -1 : 1;
                   //update currentValue + delta
                   const updates = response.entities.map((record)=> {
                      if(record.diana_sortableid!=sourceId){
                         return Xrm.WebApi.updateRecord("diana_sortable", record.diana_sortableid, {"new_task_sortorder": record.new_task_sortorder + delta});
                      }
                      return Promise.resolve();
                   })         
                   //update source mit targetValue
                   updates.push(Xrm.WebApi.updateRecord("diana_sortable", sourceId, {"new_task_sortorder": targetValue}))            
                   return Promise.all(updates).then(refreshGrid, refreshGrid)
                });           
             }
             
         
         
             
             window.addEventListener("message", (e) => { 			
                console.log("registered OnMessage", e);
                if (e.data?.messageName === "Dianamics.DragRows") {		
                  const data = e.data.data;
                  move(data.sourceId, data.sourceValue, data.targetId, data.targetValue);
                 
                 console.log(e);
                 }
              })
             }
             

/*         if (eventName) {
            const papsGridCustomizer: PAOneGridCustomizer = { cellRendererOverrides, cellEditorOverrides };
            (context as any).factory.fireEvent(eventName, papsGridCustomizer);
        } 
        
*/


        if (eventName) {
            const draggableGrid = DraggableRowsGridRenderer(context);
            const paOneGridCustomizer: PAOneGridCustomizer = { 
                cellRendererOverrides,
                cellEditorOverrides,
                gridCustomizer : draggableGrid,
                //cellCustomization : draggableGrid,
                //cellCustomization : cellRendererOverrides,
                groupSuppressAutoColumn : true,
                treeData : true
             //   cellRendererOverrides : MyCellRenderer        
            };
            (context as any).factory.fireEvent(eventName, paOneGridCustomizer, SortOrderChanged, this);            
        }  
    }

    /**
     * Called when any value in the property bag has changed. This includes field values, data-sets, global values such as container height and width, offline status, control metadata values such as label, visible, etc.
     * @param context The entire property bag available to control via Context Object; It contains values as set up by the customizer mapped to names defined in the manifest, as well as utility functions
     * @returns ReactElement root react element for the control
     */
    public updateView(context: ComponentFramework.Context<IInputs>): React.ReactElement {
        //return React.createElement(React.Fragment);
        return React.createElement(DraggableCell, {rowId: "id111", rowIndex: 1});
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
        // Add code to cleanup control if necessary
    }
}