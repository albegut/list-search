import * as React from 'react';

import styles from '../ListSearchWebPart.module.scss';
import * as strings from 'ListSearchWebPartStrings';

import { IListSearchState } from './IListSearchState';
import { IListSearchProps } from './IListSearchProps';

import { Spinner, SpinnerSize } from 'office-ui-fabric-react/lib/Spinner';
import { Label } from 'office-ui-fabric-react/lib/Label';

import IListService from '../services/IListService'
import ListService from '../services/ListService';


export default class ISecondWebPart extends React.Component<IListSearchProps, IListSearchState> {
    private listService: IListService;

    constructor(props: IListSearchProps, state: IListSearchState) {
        super(props);
        this.listService = new ListService(this.context);
        this.state = {
            items: null,
            isLoading: true,
            errorMsg: "",
        };

    }

    public componentDidMount() {
        this.readItems();
    }

    private async readItems() {
        let listItemsPromise: Promise<Array<any>>;
        let items: Array<any>;
        let viewFields: Array<string> = new Array<string>();
        viewFields.push("Title");
        viewFields.push("ID");
        try {
            items = await this.listService.getListItems(this.props.ListName, viewFields, "ID", true);
            this.setState({
                items,
                isLoading: false,
            });
        } catch (error) {
            this.setState({
                errorMsg: "readItemsError",
                isLoading: false,
            });
        }
    }

    public render(): React.ReactElement<IListSearchProps> {

        return (
            <div className={styles.listSearch}>
                <div className={styles.container}>
                    <div className={styles.row}>
                        {this.state.isLoading ? (<Spinner label="Cargando..." size={SpinnerSize.large} />) :
                            <p>The list {this.props.ListName} has {this.state.items ? this.state.items.length : 0}</p>}
                    </div>
                </div>
            </div>);
    }
}