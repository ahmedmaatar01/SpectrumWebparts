import React, { useState, useEffect, useCallback } from 'react';
import { SPListColumn, SPOperations, SPListItem } from "../../../Services/SPServices";
import styles from '../Ged365Webpart.module.scss';
import 'bootstrap/dist/css/bootstrap.min.css';
import '@fortawesome/fontawesome-free/css/all.min.css';

interface ISingleDocItemProps {
    context: any;
    column: SPListColumn;
    item: SPListItem;
    onDirectoryClick: (path: string) => void;
    text_color: string;
}

const fileIconMap: { [key: string]: string } = {
    'xlsx': 'https://res.cdn.office.net/files/fabric-cdn-prod_20240129.001/assets/item-types/20/xlsx.svg',
    'xls': 'https://res.cdn.office.net/files/fabric-cdn-prod_20240129.001/assets/item-types/20/xlsx.svg',
    'doc': 'https://res.cdn.office.net/files/fabric-cdn-prod_20240129.001/assets/item-types/20/docx.svg',
    'docx': 'https://res.cdn.office.net/files/fabric-cdn-prod_20240129.001/assets/item-types/20/docx.svg',
    'ppt': 'https://res.cdn.office.net/files/fabric-cdn-prod_20240129.001/assets/item-types/20/pptx.svg',
    'pptx': 'https://res.cdn.office.net/files/fabric-cdn-prod_20240129.001/assets/item-types/20/pptx.svg',
    'pdf': 'https://res.cdn.office.net/files/fabric-cdn-prod_20240129.001/assets/item-types/20/pdf.svg',
    'txt': 'https://res.cdn.office.net/files/fabric-cdn-prod_20240129.001/assets/item-types/20/txt.svg',
    'jpg': 'https://res.cdn.office.net/files/fabric-cdn-prod_20240129.001/assets/item-types/20/photo.svg',
    'jpeg': 'https://res.cdn.office.net/files/fabric-cdn-prod_20240129.001/assets/item-types/20/photo.svg',
    'png': 'https://res.cdn.office.net/files/fabric-cdn-prod_20240129.001/assets/item-types/20/photo.svg',
    'gif': 'https://res.cdn.office.net/files/fabric-cdn-prod_20240129.001/assets/item-types/20/photo.svg',
    'zip': 'https://res.cdn.office.net/files/fabric-cdn-prod_20240129.001/assets/item-types/20/zip.svg',
    'rar': 'https://res.cdn.office.net/files/fabric-cdn-prod_20240129.001/assets/item-types/20/zip.svg', 
    'default': 'https://res.cdn.office.net/files/fabric-cdn-prod_20240129.001/assets/item-types/20/genericfile.svg' 
};

const SingleDocItem: React.FC<ISingleDocItemProps> = ({ context, column, item, text_color, onDirectoryClick }) => {
    const [userName, setUserName] = useState<string>('');

    const fetchUserName = useCallback(async () => {
        const spOperations = new SPOperations();
        try {
            const user = await spOperations.GetUserById(context, item['responsableId']);
            setUserName(user.Title);
        } catch (error) {
            console.error(`Error fetching user with ID ${item['responsableId']}:`, error);
            setUserName('Unknown');
        }
    }, [context, item]);

    useEffect(() => {
        if (item && column.type === "User") {
            fetchUserName();
        }
    }, [item, column.type, fetchUserName]);

    const getFileIconUrl = (fileName: string): string => {
        const extension = fileName.split('.').pop()?.toLowerCase() || '';
        return fileIconMap[extension] || fileIconMap['default'];
    };

    const handleDirectoryClick = () => {
        const path = item[column.internalName];
        onDirectoryClick(path);
    };
    if (column.internalName === "FileLeafRef") {
        const fileName = item[column.internalName];
        const fileIconUrl = getFileIconUrl(fileName);

        if (item["FileSystemObjectType"] == "1") {
            return (
                <div className={styles['document-a']} onClick={handleDirectoryClick} style={{ cursor: 'pointer', color: text_color }}>
                    <img src="https://res.cdn.office.net/files/fabric-cdn-prod_20240129.001/assets/item-types/20/folder.svg" className="me-1" alt="Folder icon" />
                    {fileName}
                </div>
            );
        } else {
            const editUrl = `${item["ServerRedirectedEmbedUrl"]}&action=edit`;
            return (
                <div>     
                    <a href={editUrl} target='_blank' className={styles['document-a']} style={{  color: text_color }}>
                        <img src={fileIconUrl} className="me-1" alt="File icon" />
                        {fileName}
                    </a>
                </div>
            );
        }
    }
    if (column.type === "User") {
        return <div style={{ color: text_color }}>{userName}</div>;
    }

    return <div style={{ color: text_color }}>{item[column.internalName]}</div>;
};

export default SingleDocItem;
