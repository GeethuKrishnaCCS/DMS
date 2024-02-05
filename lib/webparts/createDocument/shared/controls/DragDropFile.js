var __extends = (this && this.__extends) || (function () {
    var extendStatics = function (d, b) {
        extendStatics = Object.setPrototypeOf ||
            ({ __proto__: [] } instanceof Array && function (d, b) { d.__proto__ = b; }) ||
            function (d, b) { for (var p in b) if (Object.prototype.hasOwnProperty.call(b, p)) d[p] = b[p]; };
        return extendStatics(d, b);
    };
    return function (d, b) {
        if (typeof b !== "function" && b !== null)
            throw new TypeError("Class extends value " + String(b) + " is not a constructor or null");
        extendStatics(d, b);
        function __() { this.constructor = d; }
        d.prototype = b === null ? Object.create(b) : (__.prototype = b.prototype, new __());
    };
})();
import * as React from 'react';
import { DragDropFiles } from "@pnp/spfx-controls-react/lib/DragDropFiles";
import { FilePicker } from '@pnp/spfx-controls-react/lib/FilePicker';
import styles from './DragDropFile.module.scss';
var DragDropFile = /** @class */ (function (_super) {
    __extends(DragDropFile, _super);
    function DragDropFile(props) {
        var _this = _super.call(this, props) || this;
        _this._getDropFiles = function (files) {
            if (files.length > 0) {
                var filename = files[0].name;
                if ((filename.substring(filename.lastIndexOf('.') + 1, filename.length) === "pdf") || (filename.substring(filename.lastIndexOf('.') + 1, filename.length) === "docx")) {
                    _this.props.returnFileData(files);
                }
                else {
                    alert("Kindly upload pdf/docx files only...");
                }
            }
        };
        _this.state = ({ files: "" });
        _this._getDropFiles = _this._getDropFiles.bind(_this);
        _this._getBrowsedFile = _this._getBrowsedFile.bind(_this);
        return _this;
    }
    DragDropFile.prototype._getBrowsedFile = function (filesReceived) {
        if (filesReceived != undefined) {
            if (filesReceived.length > 0) {
                this.props.returnFileData(filesReceived[0]);
            }
        }
    };
    DragDropFile.prototype.render = function () {
        var _this = this;
        return (React.createElement("div", { className: styles.DragDropFile },
            React.createElement(DragDropFiles, { dropEffect: "copy", enable: true, onDrop: this._getDropFiles, iconName: "Upload", labelMessage: "Drop file here..." },
                React.createElement("div", { className: styles.uploadArea },
                    React.createElement("span", null, "Drag and drop files here..."),
                    React.createElement("span", null, "OR"),
                    React.createElement(FilePicker, { accepts: [".pdf", ".docx"], buttonIcon: "FileImage", buttonLabel: "Browse", hideWebSearchTab: true, hideStockImages: true, hideOrganisationalAssetTab: true, hideOneDriveTab: true, hideSiteFilesTab: true, hideLocalMultipleUploadTab: true, hideLinkUploadTab: true, onSave: function (filePickerResult) {
                            _this._getBrowsedFile(filePickerResult);
                        }, context: this.props.context })))));
    };
    return DragDropFile;
}(React.Component));
export default DragDropFile;
//# sourceMappingURL=DragDropFile.js.map