import * as JSZip from 'jszip';
import { saveAs } from 'file-saver'


enum PartType {
    Document = "word/document.xml",
    Style = "word/styles.xml",
    Numbering = "word/numbering.xml",
    FontTable = "word/fontTable.xml",
    DocumentRelations = "word/_rels/document.xml.rels",
    NumberingRelations = "word/_rels/numbering.xml.rels",
    FontRelations = "word/_rels/fontTable.xml.rels",
}

export class OpenXmlParser {
    private zip: JSZip = new JSZip();

    fontTable: any;

    storePlaceholderChar: boolean = false;
    placeHolders: string[] = [];
    placeHolderToMerge: string = '';
    nodeTextContentToRecreate = '';
    nodesToMerge: ChildNode[] = [];

    static load(blob, data: any, targetFileName: string): void {
        var d = new OpenXmlParser();
        d.zip.loadAsync(blob).then(z => {
            var f = d.zip.files[PartType.Document];
            f ? f.async("text").then(xml => {
                let nodesToPush = [];
                var parsedDocument = new DOMParser().parseFromString(xml, "application/xml");
                if(data && data.length) {
                    d.nodeTextContentToRecreate = '';
                    d.placeHolders = [];
                    d.nodeTextContentToRecreate = '';
                    d.nodesToMerge = [];
                    d.storePlaceholderChar = false;
                    d.traverseNodes(parsedDocument.firstChild.childNodes[0].childNodes, data[0]);
                }
                if(data && data.length > 1) {
                    for(let index = 1; index < data.length; index++) {
                        d.nodeTextContentToRecreate = '';
                        d.placeHolders = [];
                        d.nodeTextContentToRecreate = '';
                        d.nodesToMerge = [];
                        d.storePlaceholderChar = false;
                        var result = new DOMParser().parseFromString(xml, "application/xml");
                        d.traverseNodes(result.firstChild.childNodes[0].childNodes, data[index]);

                        nodesToPush.push(d.getPageBreakNode())
                        result.firstChild.childNodes[0].childNodes.forEach((item, index) => {
                            if(item.nodeName !== 'w:sectPr'){
                                nodesToPush.push(item);
                            }
                        });
                    }
                }
                
                nodesToPush.map((item, index) => {
                    console.log('index => ' + index + ' <=> content => ' + item.textContent);
                    parsedDocument.firstChild.childNodes[0].lastChild.after(item)
                });
                var xmlString = new XMLSerializer().serializeToString(parsedDocument);
                d.zip.file("word/document.xml", xmlString);
                z.generateAsync({type:"blob"}).then(x=> {
                    saveAs(x, targetFileName)
                });
            }) : null;
        });
    }

    traverseNodes(cNodes: NodeListOf<ChildNode>, jsonData: any) {
        if(cNodes && cNodes.length > 0) {
            cNodes.forEach((nodeItem) => {
                if(nodeItem.nodeName ===  'w:t')
                {
                    
                    let textContent = nodeItem.textContent;
                    if(textContent) {
                        for (let i = 0; i < textContent.length; i++) {
                            if(textContent.charAt(i) === '[' && !this.storePlaceholderChar) {
                                this.storePlaceholderChar = true;
                            } else if(textContent.charAt(i) === ']' && this.storePlaceholderChar) {
                                if(this.nodesToMerge && this.nodesToMerge.length > 0) {
                                    this.nodesToMerge.push(nodeItem);
                                    this.nodeTextContentToRecreate = this.nodeTextContentToRecreate + textContent;
                                    nodeItem.textContent = this.nodeTextContentToRecreate;
                                    this.nodeTextContentToRecreate = '';
                                }
                                this.storePlaceholderChar = false;
                                this.placeHolders.push(this.placeHolderToMerge);
                                const replacedText = nodeItem.textContent.replace('[' + this.placeHolderToMerge + ']', jsonData[this.placeHolderToMerge])
                                nodeItem.textContent = replacedText;
                                this.placeHolderToMerge = '';
                            } else {
                                if(this.storePlaceholderChar) {
                                    this.placeHolderToMerge = this.placeHolderToMerge + textContent.charAt(i);
                                }
                            }
                        }
                        if(this.storePlaceholderChar) {
                            this.nodesToMerge.push(nodeItem);
                            this.nodeTextContentToRecreate = this.nodeTextContentToRecreate + textContent;
                            nodeItem.remove();
                        }
                    }
                }

                this.traverseNodes(nodeItem.childNodes, jsonData)
            })
        }
    }

    getPageBreakNode(): HTMLElement {
        let pageBreakElement = document.createElement('w:p')
        let runElement = document.createElement('w:r');
        let brElement = document.createElement('w:br')
        brElement.setAttribute('w:type', 'page');
        runElement.appendChild(brElement);
        pageBreakElement.appendChild(runElement)
        return pageBreakElement;
    }
}
