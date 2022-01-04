# json-excel

## ℹ️ About

This project was inspired by [json-as-xlsx](https://github.com/LuisEnMarroquin/json-as-xlsx), and aims to make the export of json data to Excel easier, with more *** style*** provided.

The main dependencies is [sheetjs-style](https://github.com/gitbrent/xlsx-js-style) and underlying dependencies for this project is [sheetjs](https://github.com/sheetjs/sheetjs).

All projects are under the Apache 2.0 License.

## 🔌 Installation

Install npm:

```sh
npm install json-excel --save
```

Install yarn:

```html
yarn add json-excel
```

### 🪛 Main function

* `jsonExcel(data: IJsonSheet[], settings?: ISettings)`  export excel file with multiple sheets

  - <details>
      <summary><b>Example</b> (click to show)</summary>
    
    ```js
    const baseUrl = 'http://www.baseUrl.com'
    const data = [
      {
        sheetName: 'Adults',
        header: [
          {
            label: 'ID',
            value: 'id',
            option: {
              type: 'hyperlink',
              render: row => ({
                Target: `${baseUrl}/${row?.group}/${row?.id}`,
                Tooltip: `click to visit ${baseUrl}/${row?.group}/${row?.id}`
              }),
              style: {
                fill: { fgColor: { rgb: 'FFFFAA' } },
                font: {
                  bold: true,
                  italic: true
                }
              }
            }
          },
          { label: 'User', value: 'user' }, // Top level data
          { label: 'Age', value: row => row.age + ' years' }, // Run functions
          { label: 'Phone', value: row => (row.more ? row.more.phone || '' : '') } // Deep props
        ],
        data: [
          { id: 1, group: 1, user: 'Andrea', age: 20, more: { phone: '11111111' } },
          { id: 2, group: 1, user: 'Luis', age: 21, more: { phone: '22222222' } },
          { id: 3, group: 2, user: 'Tom', age: 18, more: { phone: '33333333' } },
          { id: 4, group: 2, user: 'Jack', age: 24, more: { phone: '444444444' } }
        ]
      },
      {
        sheetName: 'Children',
        header: false,
        data: [
          { id: 11, group: 3, user: 'Manuel', age: 16, more: { phone: '55555555' } },
          { id: 12, group: 4, user: 'Ana', age: 17, more: { phone: '66666666' } }
        ]
      }
    ]
    
    const settings = {
      fileName: 'PersonalInformation',
      extraLength: 3,
    }
    
    jsonExcel(data, settings)
    ```
    
    </details>



* `jsonSheet(data: IData[], header?: IHeader[] | boolean, fileName?: string, settings?: ISettings)`  fast export single sheet excel file 


  * <details>
      <summary><b>Example</b> (click to show)</summary>

    ```js
    const baseUrl = 'http://www.baseUrl.com'
    
    const data = [
      { id: 1, group: 1, user: 'Andrea', age: 20, more: { phone: '11111111' } },
      { id: 2, group: 1, user: 'Luis', age: 21, more: { phone: '22222222' } },
      { id: 3, group: 2, user: 'Tom', age: 18, more: { phone: '33333333' } },
      { id: 4, group: 2, user: 'Jack', age: 24, more: { phone: '444444444' } }
    ]
    
    const header = [
      {
        label: 'ID',
        value: 'id',
        option: {
          type: 'hyperlink',
          render: row => ({
            Target: `${baseUrl}/${row?.group}/${row?.id}`,
            Tooltip: `click to visit ${baseUrl}/${row?.group}/${row?.id}`
          }),
          style: {
            fill: { fgColor: { rgb: 'FFFFAA' } },
            font: {
              bold: true,
              italic: true
            }
          }
        }
      },
      { label: 'User', value: 'user' }, // Top level data
      { label: 'Age', value: row => row.age + ' years' }, // Run functions
      { label: 'Phone', value: row => (row.more ? row.more.phone || '' : '') } // Deep props
    ]
    
    jsonSheet(data, header, "AdultsInformation")
    ```

    </details>



### 🗒 API

* **IJsonSheet**

| Attributes | Describe | Type                     | Required |
| ---------- | -------- | ------------------------ | -------- |
| sheetName  |          | `string`                 | `false`  |
| header     |          | `IHeader[] ` | `boolean` | `false` |
| data       |          | `IData[]`                | `true`   |

* **ISettings**

| Attributes   | Describe                                                     | Type             | Required |
| ------------ | ------------------------------------------------------------ | ---------------- | -------- |
| extraLength  | A bigger number means that columns will be wider             | `number`         | `false`  |
| fileName     | The name of the exported file                                | `string`         | `false`  |
| writeOptions | Style options from https://github.com/SheetJS/sheetjs#writing-options | `WritingOptions` | `false`  |

* **IData**

`[key: string]` : ` string | number | boolean | Date | IData | any`

* **IHeader**

| Attributes | Sub Attributes | Describe | Type                                           | Required |
| ---------- | -------------- | -------- | ---------------------------------------------- | -------- |
| label      | -              |          | `string`                                       | `true`   |
| value      | -              |          | `string`                                       | `true`   |
| option     | type           |          | `"text"` | `"hyperlink"`                       | `false`  |
|            | render         |          | `Hyperlink` | ` ((value: IData) => Hyperlink)` | `false`  |
|            | style          |          | `CellStyle`                                    | `false`  |

* **Hyperlink**

| Attributes | Describe | Type     | Required |
| ---------- | -------- | -------- | -------- |
| Target     |          | `string` | `ture`   |
| Tooltip    |          | ``string | `false`  |

* **CellStyle**

Cell styles are specified by a style object that roughly parallels the OpenXML structure. The style object has five
top-level attributes: `fill`, `font`, `numFmt`, `alignment`, and `border`.

| Style Attribute | Sub Attributes | Values                                                       |
| :-------------- | :------------- | :----------------------------------------------------------- |
| fill            | patternType    | `"solid"` or `"none"`                                        |
|                 | fgColor        | `COLOR_SPEC`                                                 |
|                 | bgColor        | `COLOR_SPEC`                                                 |
| font            | name           | `"Calibri"` // default                                       |
|                 | sz             | `"11"` // font size in points                                |
|                 | color          | `COLOR_SPEC`                                                 |
|                 | bold           | `true` or `false`                                            |
|                 | underline      | `true` or `false`                                            |
|                 | italic         | `true` or `false`                                            |
|                 | strike         | `true` or `false`                                            |
|                 | outline        | `true` or `false`                                            |
|                 | shadow         | `true` or `false`                                            |
|                 | vertAlign      | `true` or `false`                                            |
| numFmt          |                | `"0"` // integer index to built in formats, see StyleBuilder.SSF property |
|                 |                | `"0.00%"` // string matching a built-in format, see StyleBuilder.SSF |
|                 |                | `"0.0%"` // string specifying a custom format                |
|                 |                | `"0.00%;\\(0.00%\\);\\-;@"` // string specifying a custom format, escaping special characters |
|                 |                | `"m/dd/yy"` // string a date format using Excel's format notation |
| alignment       | vertical       | `"bottom"` or `"center"` or `"top"`                          |
|                 | horizontal     | `"left"` or `"center"` or `"right"`                          |
|                 | wrapText       | `true ` or ` false`                                          |
|                 | readingOrder   | `2` // for right-to-left                                     |
|                 | textRotation   | Number from `0` to `180` or `255` (default is `0`)           |
|                 |                | `90` is rotated up 90 degrees                                |
|                 |                | `45` is rotated up 45 degrees                                |
|                 |                | `135` is rotated down 45 degrees                             |
|                 |                | `180` is rotated down 180 degrees                            |
|                 |                | `255` is special, aligned vertically                         |
| border          | top            | `{ style: BORDER_STYLE, color: COLOR_SPEC }`                 |
|                 | bottom         | `{ style: BORDER_STYLE, color: COLOR_SPEC }`                 |
|                 | left           | `{ style: BORDER_STYLE, color: COLOR_SPEC }`                 |
|                 | right          | `{ style: BORDER_STYLE, color: COLOR_SPEC }`                 |
|                 | diagonal       | `{ style: BORDER_STYLE, color: COLOR_SPEC }`                 |
|                 | diagonalUp     | `true` or `false`                                            |
|                 | diagonalDown   | `true` or `false`                                            |

**COLOR_SPEC**: Colors for `fill`, `font`, and `border` are specified as objects, either:

-   `{ auto: 1}` specifying automatic values
-   `{ rgb: "FFFFAA00" }` specifying a hex ARGB value
-   `{ theme: "1", tint: "-0.25"}` specifying an integer index to a theme color and a tint value (default 0)
-   `{ indexed: 64}` default value for `fill.bgColor`

**BORDER_STYLE**: Border style is a string value which may take on one of the following values:

-   `thin`
-   `medium`
-   `thick`
-   `dotted`
-   `hair`
-   `dashed`
-   `mediumDashed`
-   `dashDot`
-   `mediumDashDot`
-   `dashDotDot`
-   `mediumDashDotDot`
-   `slantDashDot`

Borders for merged areas are specified for each cell within the merged area. So to apply a box border to a merged area of 3x3 cells, border styles would need to be specified for eight different cells:

-   left borders for the three cells on the left,
-   right borders for the cells on the right
-   top borders for the cells on the top
-   bottom borders for the cells on the left

## 🙏 Thanks

-   [sheetjs](https://github.com/SheetJS/sheetjs)
-   [json-as-xlsx](https://github.com/LuisEnMarroquin/json-as-xlsx)
-   [sheetjs-style](https://github.com/gitbrent/xlsx-js-style)

## 🔖 License

Please consult the attached LICENSE file for details. All rights not explicitly
granted by the Apache 2.0 License are reserved by the Original Author.