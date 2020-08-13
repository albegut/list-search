import * as React from 'react';

import styles from '../ListSearchWebPart.module.scss';
import * as strings from 'ListSearchWebPartStrings';

import { IListSearchState } from './IListSearchState';
import { IListSearchProps } from './IListSearchProps';

import { Spinner, SpinnerSize } from 'office-ui-fabric-react/lib/Spinner';
import { TextField } from 'office-ui-fabric-react/lib/TextField';
import { Checkbox } from 'office-ui-fabric-react/lib/Checkbox';



export default class ISecondWebPart extends React.Component<IListSearchProps, IListSearchState> {

    constructor(props: IListSearchProps, state: IListSearchState) {
        super(props);
        this.state = {
            isLoading: false,
            errorMsg: "",
        };


    }

    public render(): React.ReactElement<IListSearchProps> {

        return (
            <div className={styles.listSearch}>
                <div className={styles.container}>
                    <div className={styles.row}>
                        {this.state.isLoading ? (<Spinner label="Cargando..." size={SpinnerSize.large} />) :
                            (<div><Checkbox label="Simple" />
                                <TextField color="white" label="Propiedad" id="Property" value={this.props.description} />
                            </div>)}
                    </div>
                </div>
            </div>);
    }
}