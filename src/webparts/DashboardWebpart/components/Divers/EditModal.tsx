import React, { useState, useEffect } from 'react';
import { Modal, Button, Form } from 'react-bootstrap';
import { SPListItem, SPListColumn ,SPOperations} from "../../../Services/SPServices";

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
        if (!item) return;

        setError(null);

        const fieldsToUpdate: { [key: string]: any } = {};

        for (const key in formData) {
            if (formData[key] !== item[key]) {
                fieldsToUpdate[key] = formData[key];
            }
        }

        try {
            await new SPOperations().UpdateListItemFields(context, listTitle, item.Id, fieldsToUpdate);
            handleSave({ ...item, ...fieldsToUpdate });
            handleClose();
        } catch (error) {
            setError('Error updating item.');
            console.error(error);
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
