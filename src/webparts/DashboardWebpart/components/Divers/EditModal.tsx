import React, { useState, useEffect } from 'react';
import { Modal, Button, Form } from 'react-bootstrap';
import { SPListItem, SPListColumn, SPOperations } from "../../../Services/SPServices";

interface IEditModalProps {
    show: boolean;
    handleClose: () => void;
    item: SPListItem | null;
    columns: SPListColumn[];
    context: any;
    listTitle: string;
    handleSave: (updatedItem: SPListItem) => void;
}

const EditModal: React.FC<IEditModalProps> = ({ show, handleClose, item, columns, context, listTitle, handleSave }) => {
    const [formData, setFormData] = useState<SPListItem>({} as SPListItem);
    const [error, setError] = useState<string | null>(null);

    useEffect(() => {
        if (item) {
            setFormData(item);
        }
    }, [item]);

    const handleChange = (e: React.ChangeEvent<HTMLInputElement>) => {
        const { name, value } = e.target;
        setFormData(prevState => ({
            ...prevState,
            [name]: value
        }));
    };

    const handleSubmit = async () => {
        setError(null);
        if (item) {
            const spOperations = new SPOperations();
            try {
                await spOperations.UpdateListItem(context, listTitle, item.Id, formData);
                handleSave(formData);
                handleClose();
            } catch (error) {
                setError(error.message);
            }
        }
    };

    return (
        <Modal show={show} onHide={handleClose}>
            <Modal.Header closeButton>
                <Modal.Title>Edit File Information</Modal.Title>
            </Modal.Header>
            <Modal.Body>
                {error && <div className="alert alert-danger">{error}</div>}
                {item && (
                    <Form>
                        {columns.map((heading, idx) => (
                            <Form.Group key={idx} className="mb-3">
                                <Form.Label>{heading.title}</Form.Label>
                                <Form.Control
                                    type="text"
                                    name={heading.internalName}
                                    value={formData[heading.internalName] || ''}
                                    onChange={handleChange}
                                />
                            </Form.Group>
                        ))}
                    </Form>
                )}
            </Modal.Body>
            <Modal.Footer>
                <Button variant="secondary" onClick={handleClose}>
                    Close
                </Button>
                <Button variant="primary" onClick={handleSubmit}>
                    Save Changes
                </Button>
            </Modal.Footer>
        </Modal>
    );
};

export default EditModal;
