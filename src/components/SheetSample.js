import React from 'react';
import { Table } from './common/Table';
import { createData } from '../helpers';
import { useIbsheet } from '../hooks/ibsheetHooks';

const options = {
    Cfg: {
        SearchMode: 2,
        Alternate: 2,
        EnterMode: 5,
        ShowFilter: 1,
        Export: {
            Url: 'https://api.ibleaders.com/ibsheet/v8/'
        }
    },
    //틀고정 좌측 컬럼 설정
    LeftCols: [
        { Header: "No", Name: "SEQ", Width: "40" }
    ],
    //중앙(메인) 컬럼 설정
    Cols: [
        { Header: "Check", Type: "Bool", Name: "sCheck", Width: "60" },
        { Header: "Company", Type: "Text", Name: "sCompany", Width: "150" },
        { Header: "Country", Type: "Text", Name: "sCountry" },
        { Header: "Sale Quantity", Type: "Int", Name: "sSaleQuantity", Width: "80", EditMask: '^\\d*$' },
        { Header: "Sale Increase", Type: "Int", Name: "sSaleIncrease", Width: "80", EditMask: '^\\d*$' },
        { Header: "Price", Type: "Int", Name: "sPrice", Format: "#,### \\원", Width: "150" },
        { Header: "Satisfaction", Type: "Int", Name: "sSatisfaction", Format: "# \\%", Width: "80" },
        { Header: "Date", Type: "Date", Name: "sDate", Format: "dd-MM-yyyy", EditFormat: "dd-MM-yyyy", Width: 200 },
        { Header: "Button", Button: "Button", Type: "Button", Name: "sBtn", Width: 100, ButtonText: "Click me" },
        { Header: "Comment", Type: "Text", Name: "sComment", Width: "150", RelWidth: 1 }
    ]
}

const data = createData(100);

export const SheetSample = () => {
    const ibSheet = useIbsheet('sheet-sample');

    const onDownExcel = () => {
        ibSheet.ibsheet.down2Excel({
            fileName: 'test.xlsx'
        });
    }

    const onLoadExcel = () => {
        ibSheet.ibsheet.loadExcel({ mode: 'HeaderMatch' });
    }

    return (
        <div className="sample-sheet">
            <Table
                id="sheet-sample"
                el="sheet-sample"
                data={data}
                options={options}
            />

            <div className="bottom-section">
                <button
                    className="btn"
                    onClick={onDownExcel}
                >
                    Download excel
                </button>
                <button
                    className="btn"
                    onClick={onLoadExcel}
                >
                    Upload excel
                </button>
            </div>
        </div>
    );
};
