import * as React from 'react';
import * as ReactDom from 'react-dom';

import { Label } from 'office-ui-fabric-react/lib/Label';
import { MessageBar, MessageBarType } from 'office-ui-fabric-react/lib/MessageBar';

import styles from '../components/Holiday.module.scss';
import { CustomMessageProps } from '../components/CustomMessageProps';

export class CustomMessage extends React.Component<CustomMessageProps, any> {

    public render(): React.ReactElement<CustomMessageProps> {
        return (
            <div className={`ms-Grid-row  ${styles.row}`}>
                {(this.props.messageVisible) ?
                    <MessageBar
                        messageBarType={this.props.messageType}
                        ariaLabel='Aria help text here'
                    >
                        {this.props.messageText}
                    </MessageBar>
                    : null
                }
            </div>
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