import { SPFI } from "@pnp/sp";
import { ObjectHelper } from "../ObjectHelper";
import { IKeyValuePair } from "../Models/IKeyValuePair";

export namespace PnPjsHelper {

    export const CreateOrUpdateListItem: <T>(SP: SPFI, ListId: string, Item: { Id: number, [key: string]: any }) => Promise<T> = async <T>(SP: SPFI, ListId: string, Item: { Id: number, [key: string]: any }) => {
        let result: T = null;
        const query = SP.web.lists.getById(ListId).items;

        if (Item.Id) {
            const response = await query.getById(Item.Id).update(Item);
            result = await response.item();
        } else {
            const response = await query.add(Item);
            result = await response.data;
        }

        return ObjectHelper.ParseListItem(result);
    }

    export interface IGetListItemsAsKVPairsAdditionalProps<T> { TextField?: string, KeyField?: string, OrderByFieldoverwrite?: keyof T, additionalSelects?: string[], ODataFilter?: string }
    const DefaultGetListItemsAsKVPairsRequest: IGetListItemsAsKVPairsAdditionalProps<any> = { TextField: "Title", KeyField: "Id", additionalSelects: [], ODataFilter: "" };
    export const GetListitemsAsKVPairs: <T>(SP: SPFI, ListName: string, optionalProps?: IGetListItemsAsKVPairsAdditionalProps<T>) => Promise<IKeyValuePair[]> = async <T>(SP: SPFI, ListName: string, optionalProps: IGetListItemsAsKVPairsAdditionalProps<T>) => {
        const config = { ...DefaultGetListItemsAsKVPairsRequest, ...optionalProps };

        const result = await SP.web.lists.getByTitle(ListName).items.select(config.KeyField, config.TextField, ...config.additionalSelects).filter(config.ODataFilter).orderBy((config.OrderByFieldoverwrite || config.TextField) as string, true).top(2000)();
        const mapped = [];
        for (let item of result) {
            const kv: IKeyValuePair = { key: item[config.KeyField], text: item[config.TextField], data: {} };
            if (optionalProps.additionalSelects)
                for (let key of optionalProps.additionalSelects)
                    kv.data[key] = item[key];
            mapped.push(kv);
        }

        return mapped;
    }

    export const GetChoicesAsKVPairs: (SP: SPFI, ListName: string, FieldName: string) => Promise<IKeyValuePair[]> = async (SP: SPFI, ListName: string, FieldName: string) => {
        const result = await SP.web.lists.getByTitle(ListName).fields.getByInternalNameOrTitle(FieldName).select("Choices")();
        return result.Choices.sort().map((choice) => { return { key: choice, text: choice, data: {} } });
    }
}