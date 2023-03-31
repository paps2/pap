/* eslint-disable @typescript-eslint/no-unused-vars */
import { ContextualMenu, Label } from '@fluentui/react';
import * as React from 'react';
import { CellRendererOverrides, RECID, CellRendererProps, GetRendererParams, RowData } from '../types';
import { Icon } from "@fluentui/react/lib/Icon";
import { IInputs } from '../generated/ManifestTypes';
import { DraggableCell } from '../Controls/DraggableCell';


// interface for the row data to be sorted by the display sequence and the its parent task id
interface SortedRowData extends RowData {
  msdyn_displaysequence: null;
  msdyn_parenttaskid: null;
  msdyn_outlinelevel: number;
  children: SortedRowData[];
  isCollapsible: boolean;
  isCollapsed: boolean;
}

//Function for sorting the rows by the display sequence and the parent task id
function sortRows(rows: RowData[]): SortedRowData[] {
  const sortedRows: SortedRowData[] = [];
  const map: Map<string, SortedRowData> = new Map();

  rows.forEach(row => {
    const rowData: SortedRowData = {
      ...row,
      children: [],
      isCollapsible: false,
      isCollapsed: true,
      msdyn_displaysequence: null,
      msdyn_parenttaskid: null,
      msdyn_outlinelevel: 0
    };
    map.set(RECID, rowData);
    const parentRow = map.get((rowData as any)?.msdyn_parenttaskid);

    if (parentRow) {
      parentRow.children.push(rowData);
      parentRow.isCollapsible = true;
    } else {
      sortedRows.push(rowData);
    }
  });

  sortedRows.forEach(row => {
    row.children.sort((a, b) => (a as any).msdyn_displaysequence - (b as any).msdyn_displaysequence);
  });

  const updateRowState = (row: SortedRowData): SortedRowData => {
    if (row.children.length === 0) {
      return row;
    }

    const children = row.children.map(child => updateRowState(child));
    const isCollapsed = row.isCollapsed && children.every(child => child.isCollapsed);

    return {
      ...row,
      children,
      isCollapsed
    };
  };

  return sortedRows.map(row => updateRowState(row));
}



export const cellRendererOverrides: CellRendererOverrides = {
  

  ["Text"]: (props, col) => {
    
    const outlevel = (col.rowData as SortedRowData)?.msdyn_outlinelevel ?? 0;
    const DispSeq = (col.rowData as SortedRowData)?.msdyn_displaysequence ?? null;
    const parentRecID = (col.rowData as SortedRowData)?.msdyn_parenttaskid ?? null;
    const tasktype = (col.rowData as any)?.new_projecttasktype ?? null;
    const rowId = (col.rowData as any)?.RECID;
    const row = col.rowData as SortedRowData;

    if (col.colDefs[col.columnIndex].name === "msdyn_subject") {

      // if the cell is clicked then change chevron icon between ChevronDown or ChevronRight and expand/collapse the row
      const onCellClicked = () => {
        const row = col.rowData as SortedRowData;
        row.isCollapsed = !row.isCollapsed;
        console.log("Row sortRows:--" + row.isCollapsed);
        
      }

      if (outlevel === 1) {
        // if the record is the Project Summary Row make it bold and black
        if (tasktype === "100000004") {
          return (
            //if the icon is clicked, then the row is collapsed

            <div style={{ color: "black" }} onClick={onCellClicked}>
              <Icon iconName={row.isCollapsed ? "ChevronDown" : "ChevronRight"} />
              {"  Project: " + props.formattedValue}
            </div>

          );
        } else {
          return (
            <div style={{ color: "blue", textIndent: 10 }} onClick={onCellClicked}>
              <Icon iconName={row.isCollapsed ? "ChevronDown" : "ChevronRight"} /> {"  " + outlevel + " " + props.formattedValue}
            </div>
          );
        }
      } else {
        return (
          <Label style={{ color: "black", textIndent: 36 + outlevel * 15 }}>
            {outlevel + " " + props.formattedValue}
          </Label>
        );
      }
    } else {
      return <Label style={{ color: "green" }}>{props.formattedValue}</Label>;
    }
  },


  //export const MyCellRenderer = {
  ["Integer"]: (props, col) => {

/*     if (col.colDefs[col.columnIndex].name === 'msdyn_outlinelevel') {
      // Render the cell value in green
      if ((props.value as number) === 1) {
        props.rowHeight = 2
        return <Label style={{ color: 'blue', textAlign: 'right' }}>{props.formattedValue}</Label>
      }
      else {
        return <Label style={{ color: 'black' }}>{props.formattedValue}</Label>
      }
    } */

    const column = col.colDefs[col.columnIndex];
    if(column.name==="new_task_sortorder"){
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
       return <DraggableCell rowId={col.rowData?.[RECID]} rowIndex={props.value} text={props.formattedValue} onDropped={onDropped}/>                                
    }
    return null;
  },
//};

  

  ["OptionSet"]: (props, col) => {
    const onCellClicked = (event?: React.MouseEvent<HTMLElement, MouseEvent> | MouseEvent) => {
      if (props.startEditing) props.startEditing();
      console.log("onCellClicked----------");
    }
    return (<div onClick={onCellClicked}>
      <Icon style={{ color: 'blue' ?? "grey", marginRight: "8px" }} iconName="CircleShapeSolid" aria-hidden="true" />
      {props.formattedValue}
    </div>)
  },
  ["TwoOptions"]: (props, col) => {
    const column = col.colDefs[col.columnIndex];
    if (column.name === "msdyn_summarytask") {
      const onCellClicked = () => {
        if (props.startEditing) props.startEditing();
      }
      const smiley = props.value === "1" ? "Emoji2" : "Sad";
      const label = props.formattedValue;
      return <div onClick={onCellClicked} style={{ textAlign: "center" }}><Icon iconName={smiley} style={{ color: props.value === "1" ? "green" : "red" }}></Icon></div>
    }
  }
}