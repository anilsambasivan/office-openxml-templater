import { OpenXmlParser } from './office-openxml-parser';

export interface Options {
    inWrapper: boolean;
    ignoreWidth: boolean;
    ignoreHeight: boolean;
    ignoreFonts: boolean;
    breakPages: boolean;
    debug: boolean;
    experimental: boolean;
    className: string;
}

export function generate(file: Blob | any, data: any, targetFileName: string) {
    return OpenXmlParser.load(file, data, targetFileName);
}