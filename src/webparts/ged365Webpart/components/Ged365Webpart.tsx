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
    this.fetchDocLibColsTitles()
    this.GetDocuments();
    this.fetchFileAndFolderCounts();
    console.log("did mount >>>>>>>>>>>>>>>>>>>");
  }

  public componentDidUpdate(prevProps: IGed365WebpartProps) {
    console.log("did update >>>>>>>>>>>>>>>>>>>");

    if (prevProps.list_title !== this.props.list_title) {
      console.log("did update title changed>>>>>>>>>>>>>>>>>>>");
      this.setState({
        directory_link: this.props.context.pageContext.web.serverRelativeUrl + "/" + this.props.list_title,
      });
      console.log("this.state.directory_link--------update-------------->")
      console.log(this.state.directory_link)
      this.fetchDocLibColsTitles();
      this.GetDocuments();
      this.fetchFileAndFolderCounts();

    }
  }

  private fetchDocLibColsTitles() {
    this._spOperations
      .GetListColumns(this.props.context, this.props.list_title)
      .then((result: SPListColumn[]) => {
        const internalNames = result.map(col => col.internalName);
        this.setState({
          documents_cols: result,
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
      console.log("stats" + { fileCount, folderCount })
      console.log("files " + this.state.fileCount + "folders " + this.state.folderCount)
    } catch (error) {
      console.error('Error fetching file and folder counts:', error);
    }
  }
  private getFocLibItemsFromNavClick(clickedItem: string) {
    const indexOfClickedItem = this.state.nav_links.indexOf(clickedItem);

    if (indexOfClickedItem !== -1) {
      const updatedNavLinks = this.state.nav_links.slice(0, indexOfClickedItem + 1);
      this.setState({ nav_links: updatedNavLinks });
      //directory_link
      let new_link: string;
      new_link = this.props.context.pageContext.web.serverRelativeUrl

      updatedNavLinks.forEach(element => {
        new_link = new_link + '/' + element;
        console.log("new link :" + new_link)
        this.setState({
          directory_link: new_link,
        });
      });

      console.log("api link" + this.state.directory_link)
      this.GetDocuments();
    }
  }

  private GetDocuments() {
    this.fetchDocLibColsTitles();
    if (this.state.directory_link.indexOf(this.props.list_title) !== -1) {
      let startIndex = this.state.directory_link.indexOf(this.props.list_title);
      let nav_links_string = this.state.directory_link.substring(startIndex);
      let nav_links_tab = nav_links_string.split("/");
      console.log(nav_links_tab)

      this.setState({
        nav_links: nav_links_tab,
      });
    } else if (!this.state.directory_link || this.state.directory_link == "") {
      this.setState({
        directory_link: this.props.context.pageContext.web.serverRelativeUrl + "/" + this.props.list_title,
      });
      this.setState({
        nav_links: [this.props.list_title],
      });
    }

    console.log("GetDocuments >>>>>>>>>>>>")
    console.log("this.props.context")
    console.log(this.props.context)
    console.log("list_title >>>>>>>>>>>>")
    console.log(this.props.list_title)
    console.log("this.state.directory_link >>>>>>>>>>>>")
    console.log(this.state.directory_link)
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
    console.log("this.state.directory_link");
    console.log(this.state.directory_link);
  }

  public render(): React.ReactElement<IGed365WebpartProps> {
    const { description, hasTeamsContext } = this.props;

    if (!this.props.list_title) {
      return (
        <>
          <h4>selectionner la liste que vous souhaiter visualiser</h4>
        </>
      )
    }
    return (
      <section className={`${styles.ged365Webpart} ${hasTeamsContext ? styles.teams : ''}`}>
        <div className={styles.webpartContainer}>
          <div className='row'>

            <div className='col-2 text-white d-flex flex-column justify-content-center'>
              <h5>Bibliothéque : {this.props.list_title}</h5>
              <hr />
              <h6>Documents</h6>
              <p className={styles.kpi}>{this.state.fileCount}</p>
              <h6>Dossiers</h6>
              <p className={styles.kpi}>{this.state.folderCount}</p>
            </div>
            <div className='col-10'>
              <div className={styles['table-section']}>
                <nav aria-label="breadcrumb" >
                  <ol className={styles.breadcrumbPersonaliser}>
                    {this.state.nav_links.map((item, index) => (
                      <li key={index}><a href="#" onClick={() => this.getFocLibItemsFromNavClick(item)}>{item} </a> <i className="fas fa-chevron-right text-white p-2"></i> </li>
                    ))}
                  </ol>
                </nav>
                <TableRender
                  context={this.props.context}
                  table_headings={this.state.documents_cols}
                  table_items={this.state.listItems}
                  onDirectoryClick={this.handleDirectoryClick}
                />
              </div>
            </div>
          </div>
          <div>Description: <strong>{escape(description)}</strong></div>
        </div>
      </section>
    );
  }
}