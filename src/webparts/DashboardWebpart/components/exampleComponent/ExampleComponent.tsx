import { useState, useEffect } from 'react';
import { SPListItem ,SPOperations} from "../../../Services/SPServices";
import React from 'react';
import 'bootstrap/dist/css/bootstrap.min.css';

interface IExampleComponentProps {
    context: any;
    liste_titre: string;
    items_number: number;
}

interface IExampleComponentState {
    listItems: SPListItem[];
    selectedItem: string | null;
    showModal: boolean;
    showEditModal: boolean;
    Titre_list_item: string;
    Item_Id: string;
}

const ExampleComponent: React.FC<IExampleComponentProps> = ({ context, liste_titre, items_number }) => {
    const [state, setState] = useState<IExampleComponentState>({
        listItems: [],
        selectedItem: null,
        showModal: false,
        showEditModal: false,
        Titre_list_item: "",
        Item_Id: "",

    });


    useEffect(() => {
        setState(prevState => ({
            ...prevState,
            Titre_list_item: liste_titre
        }));
    }, [liste_titre]);


    const _spOperations = new SPOperations();
    const toggleButton = () => {
        _spOperations
            .GetDocLibItems(context, liste_titre,"")
            .then((results: SPListItem[]) => {
                setState({ listItems: results, selectedItem: null, showModal: false, showEditModal: false, Titre_list_item: "", Item_Id: "" });
                console.log("List items updated");
            })
            .catch(error => {
                console.error('Error updating list items:', error);
            });
    };


    // Render your component content here
    return (
        <div>
            <h2>Liste : {liste_titre}</h2>
            {state.Titre_list_item}
            <a className="btn btn-primary" href="#d" onClick={toggleButton}>bouton</a>

        </div>
    );
};

export default ExampleComponent;
