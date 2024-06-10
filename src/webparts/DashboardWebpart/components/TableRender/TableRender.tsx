import { useState, useEffect } from 'react';
import { SPListItem, SPListColumn } from "../../../Services/SPServices";
import React from 'react';
import 'bootstrap/dist/css/bootstrap.min.css';
import 'bootstrap/dist/js/bootstrap.min.js';
import '@fortawesome/fontawesome-free/css/all.min.css';
import SingleDocItem from '../Divers/SingleDocItem';
import styles from '../Ged365Webpart.module.scss';

interface ITableRenderProps {
    context: any;
    table_headings: SPListColumn[];
    table_items: SPListItem[];
    onDirectoryClick: (path: string) => void;
    text_color: string;
}

const TableRender: React.FC<ITableRenderProps> = ({ context, table_headings, table_items, text_color, onDirectoryClick }) => {
    const [docItems, setDocItems] = useState<SPListItem[]>([]);
    const [sortColumn, setSortColumn] = useState<string>('');
    const [sortOrder, setSortOrder] = useState<'asc' | 'desc'>('asc');

    useEffect(() => {
        setDocItems(table_items);
    }, [table_items]);

    const handleSort = (column: string) => {
        const newSortOrder = sortColumn === column && sortOrder === 'asc' ? 'desc' : 'asc';
        const sortedItems = [...docItems].sort((a, b) => {
            const aValue = a[column];
            const bValue = b[column];

            if (aValue < bValue) {
                return newSortOrder === 'asc' ? -1 : 1;
            }
            if (aValue > bValue) {
                return newSortOrder === 'asc' ? 1 : -1;
            }
            return 0;
        });

        setSortColumn(column);
        setSortOrder(newSortOrder);
        setDocItems(sortedItems);
    };

    const filteredHeadings = table_headings.filter(heading =>
        !["Title", "_ExtendedDescription", "ContentType"].includes(heading.internalName)
    );

    return (
        <div className={styles['table-section']}>
            <table className="mon-tableau">
                <thead>
                    <tr>
                        {filteredHeadings.map((heading, index) => (
                            <th key={index} onClick={() => handleSort(heading.internalName)} style={{color:text_color}}>
                                {heading.title}
                                {sortColumn === heading.internalName && (
                                    sortOrder === 'asc' ? 
                                    <i className="fas fa-chevron-up ms-1"></i> : 
                                    <i className="fas fa-chevron-down ms-1"></i>
                                )}
                            </th>
                        ))}
                        <th style={{color:text_color}}>Actions</th>
                    </tr>
                </thead>
                <tbody>
                    {docItems.map((item, index) => (
                        <tr key={index}>
                            {filteredHeadings.map((heading, idx) => (
                                <td key={idx}>
                                    <SingleDocItem
                                        context={context}
                                        column={heading}
                                        item={item}
                                        onDirectoryClick={onDirectoryClick}
                                        text_color={text_color}
                                    />
                                </td>
                            ))}
                            <td>
                                <div className="btn-group" role="group">
                                    <button id={`btnGroupDrop${index}`} type="button" className="btn btn-sm dropdown-toggle" data-bs-toggle="dropdown" aria-expanded="false" style={{  color: text_color,border:"1px solid "+ text_color }}>
                                        <i className="fas fa-ellipsis-h"></i>
                                    </button>
                                    <ul className="dropdown-menu" aria-labelledby={`btnGroupDrop${index}`}>
                                        <li><a className="dropdown-item" href="#">Edit</a></li>
                                        <li><a className="dropdown-item" href="#">Delete</a></li>
                                    </ul>
                                </div>
                            </td>
                        </tr>
                    ))}
                </tbody>
            </table>
        </div>
    );
};

export default TableRender;
