import * as React from 'react';
import { DragDropFiles } from "@pnp/spfx-controls-react/lib/DragDropFiles";
import { FilePicker, IFilePickerResult } from '@pnp/spfx-controls-react/lib/FilePicker';
import styles from './DragDropFile.module.scss';
import { IDragDropFileProps, IDragDropFileState } from '../../interfaces/IDragDropFile';

export default class DragDropFile extends React.Component<IDragDropFileProps, IDragDropFileState> {

    constructor(props: IDragDropFileProps) {
        super(props);
        this.state = ({ files: "" });
        this._getDropFiles = this._getDropFiles.bind(this);
        this._getBrowsedFile = this._getBrowsedFile.bind(this);
    }

    private _getDropFiles = (files: string | any[]): void => {
        if (files.length > 0) {
            const filename = files[0].name;
            if ((filename.substring(filename.lastIndexOf('.') + 1, filename.length) === "pdf") || (filename.substring(filename.lastIndexOf('.') + 1, filename.length) === "docx")) {
                this.props.returnFileData(files)
            }
            else {
                alert("Kindly upload pdf/docx files only...");
            }
        }
    }

    private _getBrowsedFile(filesReceived: string | any[]) {
        if (filesReceived != undefined) {
            if (filesReceived.length > 0) {
                this.props.returnFileData(filesReceived[0])
            }
        }
    }

    public render(): React.ReactElement<IDragDropFileProps> {

        return (
            <div className={styles.DragDropFile}>
                <DragDropFiles
                    dropEffect="copy"
                    enable={true}
                    onDrop={this._getDropFiles}
                    iconName="Upload"
                    labelMessage="Drop file here..."
                >
                    <div className={styles.uploadArea}>
                        <span>Drag and drop files here...</span>
                        <span>OR</span>
                        <FilePicker
                            accepts={[".pdf", ".docx"]}
                            buttonIcon="FileImage"
                            buttonLabel="Browse"
                            hideWebSearchTab={true}
                            hideStockImages={true}
                            hideOrganisationalAssetTab={true}
                            hideOneDriveTab={true}
                            hideSiteFilesTab={true}
                            hideLocalMultipleUploadTab={true}
                            hideLinkUploadTab={true}
                            onSave={(filePickerResult: IFilePickerResult[]) => {
                                this._getBrowsedFile(filePickerResult)
                            }}
                            context={this.props.context as any}
                        />
                    </div>
                </DragDropFiles>
            </div>
        );
    }
}