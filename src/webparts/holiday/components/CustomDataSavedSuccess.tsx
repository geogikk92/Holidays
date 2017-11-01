import * as React from 'react';
import * as ReactDom from 'react-dom';

import { Label } from 'office-ui-fabric-react/lib/Label';
import styles from '../components/Holiday.module.scss';
import { DefaultButton, MessageBar, MessageBarType } from 'office-ui-fabric-react';

export interface ICustomDataSavedSuccessProps {
    messageText: string;
    redirectTo: string;
    messageType: number;
}

export class CustomDataSavedSuccess extends React.Component<ICustomDataSavedSuccessProps, any> {

    public render(): React.ReactElement<ICustomDataSavedSuccessProps> {
        return (
            <MessageBar
                messageBarType={this.props.messageType}
                ariaLabel='Aria help text here'
                actions={
                    <div>
                        <DefaultButton href={this.props.redirectTo}>ОК</DefaultButton>
                    </div>
                }
            >{this.props.messageText} </MessageBar>
        );
    }
}

    // /** Info styled MessageBar */
    // info = 0,
    // /** Error styled MessageBar */
    // error = 1,
    // /** Blocked styled MessageBar */
    // blocked = 2,
    // /** SevereWarning styled MessageBar */
    // severeWarning = 3,
    // /** Success styled MessageBar */
    // success = 4,
    // /** Warning styled MessageBar */
    // warning = 5,