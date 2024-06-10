import { IDropdownOption } from "office-ui-fabric-react";
import { SPListItem,SPListColumn } from "../../Services/SPServices";
export interface IGed365WebpartState {
  listTiltes: IDropdownOption[];
  listItems: SPListItem[];
  status: string;
  Titre_list_item:string;
  showModal: boolean; // Add this property
  showEditModal: boolean; // Add this property
  listItemId:string;
  selectedFileType:string;
  documents_cols:SPListColumn[];
  items_cols: string[];
  directory_link:string;
  nav_links:string[];
  fileCount: number;
  folderCount: number;
}

