declare interface INovaPolozkaCommandSetStrings {
  Command1: string;
  Command2: string;
}

declare module 'NovaPolozkaCommandSetStrings' {
  const strings: INovaPolozkaCommandSetStrings;
  export = strings;
}
