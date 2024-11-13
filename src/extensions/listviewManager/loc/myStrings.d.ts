declare interface IListviewManagerCommandSetStrings {
  Command1: string;
  Command2: string;
}

declare module 'ListviewManagerCommandSetStrings' {
  const strings: IListviewManagerCommandSetStrings;
  export = strings;
}
