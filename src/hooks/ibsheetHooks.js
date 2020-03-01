import { useState, useEffect } from 'react';
import loader from '@ibsheet/loader';

export const useIbsheet = id => {
    const [ibsheet, setIbsheet] = useState();

    useEffect(() => {
        loader.once('created-sheet', () => {
            const sheet = window['IBSheet'];
            const sheetById = sheet.find(item => item != null && String(item.id) === String(id));
            setIbsheet(sheetById);
        });
    }, [id]);

    useEffect(() => {
        if (ibsheet) {
            // Sheet keydown event
            ibsheet.bind('onKeyDown', ({ sheet, key, event, name, prefix }) => {
                // Enter key is pressed
                if (prefix === '' && Number(key) === 13) {
                    if (Number(sheet.options.Cfg.EnterMode) === 5) {
                        const curRow = sheet.getFocusedRow();
                        const curCol = sheet.getFocusedCol();
                        const nextRow = sheet.getNextRow(curRow);
                        if (curRow) {
                            sheet.focus(nextRow, curCol);
                        }
                        return true;
                    };
                }

                return false;
            });
        }
    }, [ibsheet]);

    return { ibsheet }
}
