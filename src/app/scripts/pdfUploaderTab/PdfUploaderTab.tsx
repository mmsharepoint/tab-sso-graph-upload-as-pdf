import * as React from "react";
import { Provider, Text, Button, Header } from "@fluentui/react-northstar";
import TeamsBaseComponent, { ITeamsBaseComponentState } from "msteams-react-base-component";
import * as microsoftTeams from "@microsoft/teams-js";
import Axios from "axios";
import jwt_decode from "jwt-decode";
/**
 * State for the pdfUploaderTabTab React component
 */
export interface IPdfUploaderTabState extends ITeamsBaseComponentState {
    entityId?: string;
    name?: string;
    error?: string;
    token: string;
    highlight: boolean;
}

/**
 * Properties for the pdfUploaderTabTab React component
 */
export interface IPdfUploaderTabProps {

}

/**
 * Implementation of the PDF Uploader content page
 */
export class PdfUploaderTab extends TeamsBaseComponent<IPdfUploaderTabProps, IPdfUploaderTabState> {

    public async componentWillMount() {
        this.updateTheme(this.getQueryVariable("theme"));


        microsoftTeams.initialize(() => {
            microsoftTeams.registerOnThemeChangeHandler(this.updateTheme);
            microsoftTeams.getContext((context) => {
                this.setState({
                    entityId: context.entityId
                });
                this.updateTheme(context.theme);
                microsoftTeams.authentication.getAuthToken({
                    successCallback: (token: string) => {
                        this.setState({ token: token });                        
                        microsoftTeams.appInitialization.notifySuccess();
                    },
                    failureCallback: (message: string) => {
                        this.setState({ error: message });
                        microsoftTeams.appInitialization.notifyFailure({
                            reason: microsoftTeams.appInitialization.FailedReason.AuthFailed,
                            message
                        });
                    },
                    resources: [process.env.PDFUPLOADER_APP_URI as string]
                });
            });
        });
    }

    private allowDrop = (event) => {
        event.preventDefault();
        event.stopPropagation();
        event.dataTransfer.dropEffect = 'copy';
    }
    private enableHighlight = (event) => {
        this.allowDrop(event);
        this.setState({
            highlight: true
        });
    }
    private disableHighlight = (event) => {
        this.allowDrop(event);
        this.setState({
            highlight: false
        });
    }
    private dropFile = (event) => {
        this.allowDrop(event);
        let dt = event.dataTransfer;
        let files =  Array.prototype.slice.call(dt.files); //[...dt.files];
        files.forEach(fileToUpload => {
        //if (Utilities.validFileExtension(fileToUpload.name)) {
            this.uploadFile(fileToUpload);
        //}      
        });
        
    }
    private uploadFile = (fileToUpload: File) => {
        let formData = new FormData();
        formData.append('file', fileToUpload);
        formData.append('fileName', fileToUpload.name);
        formData.append('fileType', fileToUpload.type);
        Axios.post(`https://${process.env.HOSTNAME}/api/upload`, formData, {
                            headers: {
                                'Authorization': `Bearer ${this.state.token}`,
                                'content-type': 'multipart/form-data'
                            }
                        }).then(result => {
                            console.log(result);
                        });
    }
    /**
     * The render() method to create the UI of the tab
     */
    public render() {
        return (
            <Provider theme={this.state.theme}>
                <div className='dropZoneBG'>
                    Drag your file here:
                    <div className={ `dropZone ${this.state.highlight ? 'dropZoneHighlight' : ''}` }
                            onDragEnter={this.enableHighlight}
                            onDragLeave={this.disableHighlight}
                            onDragOver={this.allowDrop}
                            onDrop={this.dropFile}>
                        <div className='inner'>
                        </div>
                    </div>
                </div>
            </Provider>
        );
    }
}
