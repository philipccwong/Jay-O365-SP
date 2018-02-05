//SP Internal Field Mapping 
export interface InternalKey
{
    key: string;
    value: string;
}

export interface InternalKeys
{
    value: InternalKey[];
}

export default class InternalValue{

    public static _items: InternalKey[] 
    = [{key: '~', value: '_x007e_'},
    {key: '!', value: '_x0021_'},
    {key: '@', value: '_x0040_'},
    {key: '#', value: '_x0023_'},
    {key: '$', value: '_x0024_'},
    {key: '%', value: '_x0025_'},
    {key: '^', value: '_x005e_'},
    {key: '&', value: '_x0026_'},
    {key: '*', value: '_x002a_'},
    {key: '(', value: '_x0028_'},
    {key: ')', value: '_x0029_'},
    {key: '+', value: '_x002b_'},
    {key: '–', value: '_x002d_'},
    {key: '=', value: '_x003d_'},
    {key: '{', value: '_x007b_'},
    {key: '}', value: '_x007d_'},
    {key: ':', value: '_x003a_'},
    {key: '“', value: '_x0022_'},
    {key: '|', value: '_x007c_'},
    {key: ';', value: '_x003b_'},
    {key: '‘', value: '_x0027_'},
    {key: `\\`, value: '_x005c_'},
    {key: '<', value: '_x003c_'},
    {key: '>', value: '_x003e_'},
    {key: '?', value: '_x003f_'},
    {key: ',', value: '_x002c_'},
    {key: '.', value: '_x002e_'},
    {key: '/', value: '_x002f_'},
    {key: '`', value: '_x0060_'},
    {key: ' ', value: '_x0020_'}
    ];

    public static get(): Promise<InternalKey[]> {
        return new Promise<InternalKey[]>((resolve) => {
                resolve(InternalValue._items);
            });
        }

}
