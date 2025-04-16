[**aobajsxl v1.0.0**](README.md)

***

[aobajsxl](globals.md) / Excel

# Class: Excel

Defined in: [index.ts:19](https://github.com/jchirayil/aobajsxl/blob/3b0f17941fd9e229adc6897ac47ec47bc88f8d23/src/index.ts#L19)

Excel class to handle reading and writing Excel files.

## Example

```ts
const excel = new Excel();
await excel.read('example.xlsx');
const data = excel.getData('Sheet1');
console.log(data);
await excel.write('output.xlsx');
const sheetNames = excel.getSheetNames();
console.log(sheetNames);
```

## Constructors

### Constructor

> **new Excel**(): `Excel`

Defined in: [index.ts:23](https://github.com/jchirayil/aobajsxl/blob/3b0f17941fd9e229adc6897ac47ec47bc88f8d23/src/index.ts#L23)

#### Returns

`Excel`

## Methods

### getData()

> **getData**(`sheetName`): `any`[]

Defined in: [index.ts:84](https://github.com/jchirayil/aobajsxl/blob/3b0f17941fd9e229adc6897ac47ec47bc88f8d23/src/index.ts#L84)

Gets the data from a specific sheet.

#### Parameters

##### sheetName

`string`

The name of the sheet to get data from.

#### Returns

`any`[]

An array of JSON data from the specified sheet.

#### Example

```ts
const excel = new Excel();
await excel.read('example.xlsx');
const data = excel.getData('Sheet1');
console.log(data);
```

***

### getSheetNames()

> **getSheetNames**(): `string`[]

Defined in: [index.ts:69](https://github.com/jchirayil/aobajsxl/blob/3b0f17941fd9e229adc6897ac47ec47bc88f8d23/src/index.ts#L69)

Gets the names of all sheets in the Excel file.

#### Returns

`string`[]

An array of sheet names.

#### Example

```ts
const excel = new Excel();
await excel.read('example.xlsx');
const sheetNames = excel.getSheetNames();
console.log(sheetNames);
```

***

### read()

> **read**(`fileName`): `Promise`\<`void`\>

Defined in: [index.ts:39](https://github.com/jchirayil/aobajsxl/blob/3b0f17941fd9e229adc6897ac47ec47bc88f8d23/src/index.ts#L39)

Reads an Excel file and parses its data.

#### Parameters

##### fileName

`string`

The name of the file to read.

#### Returns

`Promise`\<`void`\>

A promise that resolves when the file is read.

#### Example

```ts
const excel = new Excel();
await excel.read('example.xlsx');
const data = excel.getData('Sheet1');
console.log(data);
```

***

### setData()

> **setData**(`sheetName`, `data`): `void`

Defined in: [index.ts:97](https://github.com/jchirayil/aobajsxl/blob/3b0f17941fd9e229adc6897ac47ec47bc88f8d23/src/index.ts#L97)

Sets the data for a specific sheet.

#### Parameters

##### sheetName

`string`

The name of the sheet to set data for.

##### data

`any`[]

An array of JSON data to set in the specified sheet.

#### Returns

`void`

#### Example

```ts
const excel = new Excel();
excel.setData('TestSheet', [{ name: 'Alice', age: 30 },{ name: 'Bob', age: 25 }]);
```

***

### write()

> **write**(`filename`): `Promise`\<`void`\>

Defined in: [index.ts:54](https://github.com/jchirayil/aobajsxl/blob/3b0f17941fd9e229adc6897ac47ec47bc88f8d23/src/index.ts#L54)

Writes the data to an Excel file.

#### Parameters

##### filename

`string`

The name of the file to write.

#### Returns

`Promise`\<`void`\>

A promise that resolves when the file is written.

#### Example

```ts
const excel = new Excel();
excel.setData('TestSheet', [{ name: 'Alice', age: 30 },{ name: 'Bob', age: 25 }]);
await excel.write('output.xlsx');
```
