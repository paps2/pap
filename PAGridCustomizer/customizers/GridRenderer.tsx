import { GetHeaderParams, GetRendererParams, GridCustomizer, NoRowsOverlayConfig, RECID } from "../types";
import * as React from 'react';
import { DraggableCell } from "../Controls/DraggableCell";
import { IInputs } from "../generated/ManifestTypes";


/*
export const cellEditorOverrides: CellEditorOverrides = {
    ["Text"]: (defaultProps: CellEditorProps, rendererParams: GetEditorParams) => {
      // TODO: Add your custom cell editor overrides here
      return null
    }
}*/


/* export class DianaGridRenderer implements GridCustomizer {
    GetLoadingRowRenderer(): React.ReactElement{
        return <div>Loading...</div>;
    }  
    GetCellRenderer(params: GetRendererParams): React.ReactElement{
        return <div>Test1</div>
    }
    GetHeaderRenderer(params: GetHeaderParams): React.ReactElement{
        return <div>HEADER</div>;
    }
} */

export const DraggableRowsGridRenderer = (context: ComponentFramework.Context<IInputs>) : GridCustomizer  => {    
    
    return {
    GetLoadingRowRenderer: (): React.ReactElement => {
        return <div>Loading...</div>;
    } , 
    GetCellRenderer : (params: GetRendererParams): React.ReactElement => {
        
        const cellName = params.colDefs[params.columnIndex].name;
        const formattedValue = params.colDefs[params.columnIndex].getFormattedValue(params.rowData?.[RECID]) ?? (params.rowData as any)[cellName];
        if(cellName==="new_task_sortorder"){
           
            const index = (params.rowData as any)?.new_task_sortorder ?? (params as any).rowIndex;        
            const onDropped = (sourceId: string, sourceValue:number, targetId : string, targetValue:number) => {                              
                Array.from(parent.frames).forEach((frame) => {
                    frame.postMessage({
                        messageName: "Dianamics.DragRows", 
                        data: {
                            sourceId, 
                            sourceValue,
                            targetId, 
                            targetValue
                        }
                    }, "*");
                });    
                                    
            }
            return (<div>
                <DraggableCell rowId={params.rowData?.[RECID]} rowIndex={index} text={formattedValue} onDropped={onDropped}/>                         
            </div>)
        }
        else{
            return (<div>                
                {formattedValue}
            </div>)
        }                
    },
   GetHeaderRenderer : (params: GetHeaderParams): React.ReactElement => {
        const colDisplayName = params.colDefs[params.columnIndex].displayName;
        return <div>{colDisplayName}</div>;
    }
}}