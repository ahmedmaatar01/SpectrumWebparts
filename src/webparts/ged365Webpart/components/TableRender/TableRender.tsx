import { useState, useEffect } from 'react';
import { SPListItem, SPListColumn } from "../../../Services/SPServices";
import React from 'react';
import 'bootstrap/dist/css/bootstrap.min.css';
import '@fortawesome/fontawesome-free/css/all.min.css';
import SingleDocItem from '../Divers/SingleDocItem';
import styles from '../Ged365Webpart.module.scss';

interface ITableRenderProps {
    context: any;
    table_headings: SPListColumn[];
    table_items: SPListItem[];
    onDirectoryClick: (path: string) => void;
}

interface ITableRenderState {
    doc_items: SPListItem[];
}

const TableRender: React.FC<ITableRenderProps> = ({ context, table_headings, table_items, onDirectoryClick }) => {
    const [state, setState] = useState<ITableRenderState>({
        doc_items: [],
    });

    // Ensure table_headings and table_items are arrays
    const validTableHeadings = Array.isArray(table_headings) ? table_headings : [];
    const validTableItems = Array.isArray(table_items) ? table_items : [];

    const filteredHeadings = validTableHeadings.filter(heading =>
        !["Title", "_ExtendedDescription", "ContentType"].includes(heading.internalName)
    );

    useEffect(() => {
        setState(prevState => ({
            ...prevState,
            doc_items: validTableItems
        }));
    }, [validTableItems, validTableHeadings, table_items]);

    return (
        <>
            <div className={styles['table-section']}>
                <div className='table-responsive'>


                    <table className="mon-tableau">
                        <thead>
                            <tr>
                                {filteredHeadings.map((heading, index) => (
                                    <th key={index}>{heading.title}</th>
                                ))}
                                <th>Actions</th>
                            </tr>
                        </thead>
                        <tbody>
                            {state.doc_items.map((item, index) => (
                                <tr key={index}>
                                    {filteredHeadings.map((heading, idx) => (
                                        <td key={idx}>
                                            <SingleDocItem
                                                context={context}
                                                column={heading}
                                                item={item}
                                                onDirectoryClick={onDirectoryClick}
                                            />
                                        </td>
                                    ))}
                                    <td>
                                        <div className="btn-group" role="group">
                                            <button id="btnGroupDrop1" type="button" className="btn btn-sm btn-secondary dropdown-toggle" data-toggle="dropdown" aria-haspopup="true" aria-expanded="false">
                                                <i className="fas fa-ellipsis-h"></i>
                                            </button>
                                            <div className="dropdown-menu" aria-labelledby="btnGroupDrop1">
                                                <a className="dropdown-item" href="#">edit</a>
                                                <a className="dropdown-item" href="#">delete</a>
                                            </div>
                                        </div>
                                    </td>
                                </tr>
                            ))}
                        </tbody>
                    </table>
                </div>
            </div>
        </>
    );
};

export default TableRender;
