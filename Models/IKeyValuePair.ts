export interface IKeyValuePair {
    key: number | string;
    text: string;
    data?: { [key: string]: any }
}