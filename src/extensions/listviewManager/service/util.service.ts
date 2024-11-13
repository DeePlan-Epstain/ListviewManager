export const Storage = {
    saveToStorage(key: string, value: any) {
        sessionStorage.setItem(key, JSON.stringify(value))
    },
    loadFromStorage(key: string) {
        // return JSON.parse(sessionStorage.getItem(key))
    }
}

export function makeId(length: number = 11, isNumberOnly?: boolean): string {
    let txt = ''
    let txtPossible = 'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789'
    let numberPossible = '0123456789'
    for (let i = 0; i < length; i++) {
        if (isNumberOnly) txt += numberPossible.charAt(Math.floor(Math.random() * numberPossible.length))
        else txt += txtPossible.charAt(Math.floor(Math.random() * txtPossible.length))
    }
    return txt;
}

export const sortItems = {
    number: (Items: NonNullable<any>, Field: NonNullable<string>, IsDescending?: boolean): any[] => {
        if (!Items?.length) return [];
        return Items.sort((A: any, B: any) => IsDescending ? A[Field] - B[Field] : B[Field] - A[Field])
    },
    text: (Items: NonNullable<any>, Field: NonNullable<string>, IsDescending?: boolean): any[] => {
        if (!Items?.length) return [];
        return Items.sort((A: any, B: any) => IsDescending ? B[Field].localeCompare(A[Field]) : A[Field].localeCompare(B[Field]))
    },
    date: (Items: NonNullable<any>, Field: NonNullable<string>, IsDescending?: boolean): any[] => {
        if (!Items?.length) return [];
        return Items.sort((A: any, B: any) => IsDescending ? (new Date(A[Field]).getTime() - new Date(B[Field]).getTime()) : (new Date(B[Field]).getTime() - new Date(A[Field]).getTime()));
    }
}

export function convertToSpDate(ReleventDate: any): string {
    // Get day,month and year
    let dd = String(ReleventDate.getDate());
    let mm = String(ReleventDate.getMonth() + 1); //January is 0!
    let yyyy = String(ReleventDate.getFullYear());
    if (parseInt(dd) < 10) dd = '0' + dd;
    if (parseInt(mm) < 10) mm = '0' + mm;
    // Create sp date
    let FormattedReleventDate = yyyy + '-' + mm + '-' + dd + 'T00:00:00Z';
    return FormattedReleventDate;
}

export function decimalToBinaryArray(decimal: number) {
    // Step 1: Convert the decimal number to a binary string
    const binaryString = decimal.toString(2);

    // Step 2: Split the binary string into an array of characters
    const binaryArray = binaryString.split('');

    // Step 3: Convert each character to a number
    const resultArray = binaryArray.map(bit => parseInt(bit, 10));

    return resultArray;
}