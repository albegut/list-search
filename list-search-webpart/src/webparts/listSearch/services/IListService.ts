
export default interface IListService {
    getListItems(listName:string, fields: Array<string>, orderBy:string, asc: boolean): Promise<Array<any>>
}