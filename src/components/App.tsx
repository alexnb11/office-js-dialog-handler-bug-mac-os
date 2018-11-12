import * as React from 'react';
import { Button } from 'office-ui-fabric-react';
import Progress from './Progress';

export interface AppProps {
    title: string;
    isOfficeInitialized: boolean;
}

export interface ResultList {
    resultItems: string[]
}

export default class App extends React.Component<AppProps, ResultList> {
    constructor(props, context) {
        super(props, context);
        this.openDialog = this.openDialog.bind(this);
        this.dialogEventHandler = this.dialogEventHandler.bind(this);
        this.state = {
            resultItems: []
        };
    }

    click = () => {
        Office.context.ui.displayDialogAsync('https://localhost:3000/function-file/function-file.html', {height: 50, width: 50}, this.openDialog);
    }

    openDialog = (asyncResult) => {
        const dialog = asyncResult.value;
        dialog.addEventHandler(Office.EventType.DialogMessageReceived, (arg) => { this.dialogEventHandler(arg, dialog) });
    }

    dialogEventHandler = (arg, dialog) => {
        const { resultItems } = this.state;

        this.setState({resultItems: [...resultItems, arg.message]});

        if (arg.message === 'closeDialog') {
            dialog.close();
        }
    }

    clearResult = () => {
        this.setState({resultItems: []});
    }

    render() {
        const {
            title,
            isOfficeInitialized,
        } = this.props;

        const { resultItems } = this.state;

        if (!isOfficeInitialized) {
            return (
                <Progress
                    title={title}
                    logo='assets/logo-filled.png'
                    message='Please sideload your addin to see app body.'
                />
            );
        }

        const resultListItems = resultItems.map((item, index) => (
            <div>{`${index}: Handling message id ${item}`} </div>
        ));

        return (
            <div className='ms-welcome'>
                <div>
                    <Button className='open-button' onClick={this.click}>Open dialog</Button>
                </div>
                <div>RESULT</div>
                <div className='result-box'>
                    {resultListItems}
                </div>
                <div>
                    <Button className='clear-result-button' onClick={this.clearResult}>Clear result</Button>
                </div>
            </div>
        );
    }
}
