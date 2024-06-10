// Main Component (Ged365Webpart)

import * as React from 'react';
import styles from './Ged365Webpart.module.scss';
import type { IGed365WebpartProps } from './IGed365WebpartProps';
import { IGed365WebpartState } from './IGed365WebpartState';
import { SPOperations, SPListColumn } from "../../Services/SPServices";
import { escape } from '@microsoft/sp-lodash-subset';
import TableRender from './TableRender/TableRender';

export default class Ged365Webpart extends React.Component<
  IGed365WebpartProps,
  IGed365WebpartState,
  { showModal: boolean }
> {

  public _spOperations: SPOperations;
  public slectedlisttitle: string;

  constructor(props: IGed365WebpartProps) {
    super(props);
    this._spOperations = new SPOperations();
    this.state = {
      listTiltes: [],
      listItems: [],
      status: "",
      Titre_list_item: "",
      selectedFileType: "txt",
      showModal: false,
      showEditModal: false,
      listItemId: "",
      documents_cols: [],
      items_cols: [],
      directory_link: "",
      nav_links: [],
      fileCount: 0,
      folderCount: 0,
    };
  }

  public componentDidMount() {
    this.fetchDocLibColsTitles();
    this.GetDocuments();
    this.fetchFileAndFolderCounts();
  }

  public componentDidUpdate(prevProps: IGed365WebpartProps) {
    if (prevProps.list_title !== this.props.list_title || prevProps.selectedColumns !== this.props.selectedColumns) {
      this.setState({
        directory_link: this.props.context.pageContext.web.serverRelativeUrl + "/" + this.props.list_title,
      }, () => {
        this.fetchDocLibColsTitles();
        this.GetDocuments();
        this.fetchFileAndFolderCounts();
      });
    }
  }

  private fetchDocLibColsTitles() {
    this._spOperations
      .GetListColumns(this.props.context, this.props.list_title)
      .then((result: SPListColumn[]) => {
        const internalNames = result.map(col => col.internalName);
        const filteredCols = result.filter(col => this.props.selectedColumns.includes(col.internalName));
        this.setState({
          documents_cols: filteredCols,
          items_cols: internalNames
        });
      })
      .catch(error => {
        console.error('Error fetching doc cols:', error);
      });
  }

  private async fetchFileAndFolderCounts() {
    try {
      const { fileCount, folderCount } = await this._spOperations.GetFileAndFolderCounts(this.props.context, this.props.list_title);
      this.setState({ fileCount, folderCount });
    } catch (error) {
      console.error('Error fetching file and folder counts:', error);
    }
  }

  private getFocLibItemsFromNavClick(clickedItem: string) {
    const indexOfClickedItem = this.state.nav_links.indexOf(clickedItem);

    if (indexOfClickedItem !== -1) {
      const updatedNavLinks = this.state.nav_links.slice(0, indexOfClickedItem + 1);
      this.setState({ nav_links: updatedNavLinks });
      let new_link: string = this.props.context.pageContext.web.serverRelativeUrl;

      updatedNavLinks.forEach(element => {
        new_link = new_link + '/' + element;
        this.setState({
          directory_link: new_link,
        });
      });

      this.GetDocuments();
    }
  }

  private GetDocuments() {
    if (this.state.directory_link.indexOf(this.props.list_title) !== -1) {
      const startIndex = this.state.directory_link.indexOf(this.props.list_title);
      const nav_links_string = this.state.directory_link.substring(startIndex);
      const nav_links_tab = nav_links_string.split("/");

      this.setState({
        nav_links: nav_links_tab,
      });
    } else if (!this.state.directory_link || this.state.directory_link === "") {
      this.setState({
        directory_link: this.props.context.pageContext.web.serverRelativeUrl + "/" + this.props.list_title,
        nav_links: [this.props.list_title],
      });
    }

    if (!this.state.directory_link || this.state.directory_link.indexOf(this.props.list_title) === -1) {
      this.setState({
        directory_link: this.props.context.pageContext.web.serverRelativeUrl + "/" + this.props.list_title,
        nav_links: [this.props.list_title]
      }, () => {
        this._spOperations
          .GetDocLibItems(this.props.context, this.props.list_title, this.state.directory_link)
          .then((results: any[]) => {
            this.setState({ listItems: results });
          })
          .catch(error => {
            console.error('Error updating list items:', error);
          });
      });
    } else {
      this._spOperations
        .GetDocLibItems(this.props.context, this.props.list_title, this.state.directory_link)
        .then((results: any[]) => {
          this.setState({ listItems: results });
        })
        .catch(error => {
          console.error('Error updating list items:', error);
        });
    }
  }

  private handleDirectoryClick = (path: string) => {
    this.setState({ directory_link: this.state.directory_link + "/" + path }, () => {
      this.GetDocuments();
    });
  }

  public render(): React.ReactElement<IGed365WebpartProps> {
    const { description, hasTeamsContext } = this.props;

    if (!this.props.list_title) {
      return (
        <>
          <h4>Selectionner la liste que vous souhaiter visualiser</h4>
        </>
      )
    }

    return (
      <section className={`${styles.ged365Webpart} ${hasTeamsContext ? styles.teams : ''}`}>
        <div className={styles.webpartContainer} style={{ backgroundColor: this.props.backgroundColor }}>
          <div className='row'>
            <div className='col-2 text-white d-flex flex-column justify-content-center' >
              <h5 style={{color:this.props.textColor}}>Bibliothéque : {this.props.list_title}</h5>
              <hr style={{color:this.props.textColor}} />
              <h6 style={{color:this.props.textColor}}>Documents</h6>
              <p className={styles.kpi} style={{ WebkitTextStroke:0.5,WebkitTextStrokeColor:this.props.textColor,color:'transparent'}}>{this.state.fileCount}</p>
              <h6 style={{color:this.props.textColor}}>Dossiers</h6>
              <p className={styles.kpi} style={{ WebkitTextStroke:0.5,WebkitTextStrokeColor:this.props.textColor,color:'transparent'}}>{this.state.folderCount}</p>
            </div>
            <div className='col-10'>
              <div className={styles['table-section']}>
                <nav aria-label="breadcrumb">
                  <ol className={styles.breadcrumbPersonaliser}>
                    {this.state.nav_links.map((item, index) => (
                      <li key={index}><a href="#" style={{ color: this.props.textColor }} onClick={() => this.getFocLibItemsFromNavClick(item)}>{item} </a> <i style={{ color: this.props.textColor }} className="fas fa-chevron-right  p-2"></i> </li>
                    ))}
                  </ol>
                </nav>
                <TableRender
                  context={this.props.context}
                  table_headings={this.state.documents_cols}
                  table_items={this.state.listItems}
                  onDirectoryClick={this.handleDirectoryClick}
                  text_color={this.props.textColor}
                  listTitle={this.props.list_title}
                />
              </div>
            </div>
          </div>
          <div style={{ color: this.props.textColor }}> <strong>{escape(description)}</strong></div>
        </div>
      </section>
    );
  }
}
