import { useState, useEffect } from 'react';
import { SPListItem, SPListColumn, SPOperations } from "../../../Services/SPServices";
import React from 'react';
import 'bootstrap/dist/css/bootstrap.min.css';
import 'bootstrap/dist/js/bootstrap.min.js';
import '@fortawesome/fontawesome-free/css/all.min.css';
import SingleDocItem from '../Divers/SingleDocItem';
import EditModal from '../Divers/EditModal';
import { Modal, Button } from 'react-bootstrap';
import styles from '../Ged365Webpart.module.scss';

interface ITableRenderProps {
    context: any;
    table_headings: SPListColumn[];
    table_items: SPListItem[];
    onDirectoryClick: (path: string) => void;
    text_color: string;
    listTitle: string;
}

const TableRender: React.FC<ITableRenderProps> = ({ context, table_headings, table_items, text_color, onDirectoryClick, listTitle }) => {
    const [docItems, setDocItems] = useState<SPListItem[]>([]);
    const [sortColumn, setSortColumn] = useState<string>('');
    const [sortOrder, setSortOrder] = useState<'asc' | 'desc'>('asc');
    const [showEditModal, setShowEditModal] = useState(false);
    const [showDeleteModal, setShowDeleteModal] = useState(false);
    const [selectedItem, setSelectedItem] = useState<SPListItem | null>(null);

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

    const handleEditClick = (item: SPListItem) => {
        setSelectedItem(item);
        setShowEditModal(true);
    };

    const handleDeleteClick = (item: SPListItem) => {
        setSelectedItem(item);
        setShowDeleteModal(true);
    };

    const handleCloseEditModal = () => setShowEditModal(false);
    const handleCloseDeleteModal = () => setShowDeleteModal(false);

    const handleSave = (updatedItem: SPListItem) => {
        setDocItems(prevItems => 
            prevItems.map(item => item.Id === updatedItem.Id ? updatedItem : item)
        );
    };

    const confirmDelete = async () => {
        if (!selectedItem) return;

        try {
            await new SPOperations().DeleteListItem(context, listTitle, selectedItem.Id);
            setDocItems(prevItems => prevItems.filter(item => item.Id !== selectedItem.Id));
        } catch (error) {
            console.error('Error deleting item:', error);
        } finally {
            setShowDeleteModal(false);
        }
    };

    return (
        <div className={styles['table-section']}>
            <div className='table-responsive'>
                <table className="mon-tableau">
                    <thead>
                        <tr>
                            {filteredHeadings.map((heading, index) => (
                                <th key={index} onClick={() => handleSort(heading.internalName)} style={{ color: text_color }}>
                                    {heading.title}
                                    {sortColumn === heading.internalName && (
                                        sortOrder === 'asc' ?
                                            <i className="fas fa-chevron-up ms-1"></i> :
                                            <i className="fas fa-chevron-down ms-1"></i>
                                    )}
                                </th>
                            ))}
                            <th style={{ color: text_color }}>Actions</th>
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
                                    <a style={{ color: text_color }} href="#" onClick={() => handleEditClick(item)}><i className="fa-solid fa-pen-to-square me-2"></i></a>
                                    <a style={{ color: text_color }} href="#" onClick={() => handleDeleteClick(item)}><i className="fa-solid fa-trash"></i></a>
                                </td>
                            </tr>
                        ))}
                    </tbody>
                </table>
            </div>

            <EditModal
                show={showEditModal}
                handleClose={handleCloseEditModal}
                item={selectedItem}
                columns={filteredHeadings}
                context={context}
                listTitle={listTitle}
                handleSave={handleSave}
            />

            <Modal show={showDeleteModal} onHide={handleCloseDeleteModal}>
                <Modal.Header closeButton>
                    <Modal.Title>Confirm Delete</Modal.Title>
                </Modal.Header>
                <Modal.Body>
                    Are you sure you want to delete "{selectedItem?.FileLeafRef}"?
                </Modal.Body>
                <Modal.Footer>
                    <Button variant="secondary" onClick={handleCloseDeleteModal}>
                        No
                    </Button>
                    <Button variant="danger" onClick={confirmDelete}>
                        Yes
                    </Button>
                </Modal.Footer>
            </Modal>
        </div>
    );
};

export default TableRender;