import { ListViewCommandSetContext } from "@microsoft/sp-listview-extensibility";
export class DnDService {
    private renderFillingModal: (foldersMap: Map<string, File[]>) => void;

    constructor(renderFillingModal: (foldersMap: Map<string, File[]>) => void) {
        this.renderFillingModal = renderFillingModal;
    }

    public startDnDBlock(isSecondTime?: boolean) {
        const dropZone = document.querySelectorAll("[role=presentation]");
        const testArr: Element[] = [];

        dropZone.forEach((dz) => {
            if (
                dz.className.includes("root") &&
                dz.className.includes("absolute") &&
                dz.attributes.getNamedItem("data-drop-target-key")
            )
                testArr.push(dz);
        });

        if (testArr.length) {
            testArr[0].addEventListener("drop", (ev: DragEvent) => this.handleDrop(ev));
        } else if (!isSecondTime) setTimeout(() => this.startDnDBlock(true), 100);
        else {
            console.warn("First DnD Block not found");
            this.startSecondDnDBlock();
        }
    }

    private startSecondDnDBlock(isSecondTime?: boolean) {
        let dropZone = document.querySelectorAll('[data-automationid="main"]');

        if (dropZone?.length) dropZone[0].addEventListener("drop", (ev: DragEvent) => this.handleDrop(ev));
        else if (!isSecondTime) this.startSecondDnDBlock(true);
        else console.warn("Second DnD Block not found");
    }

    private async handleDrop(event: DragEvent) {
        console.clear();
        event.stopImmediatePropagation();
        event.preventDefault();

        const foldersMap = new Map<string, File[]>();
        const promises: Promise<void>[] = [];

        for (let i = 0; i < event.dataTransfer.items.length; i++) {
            const entry = event.dataTransfer.items[i].webkitGetAsEntry();

            if (!entry) continue;

            if (entry.isDirectory) {
                promises.push((async () => {
                    const result = await this.readDirectory(entry);
                    result.forEach((files, path) => {
                        foldersMap.set(path, [...(foldersMap.get(path) || []), ...files]);
                    });
                })());
            } else if (entry.isFile) {
                const file = event.dataTransfer.items[i].getAsFile();

                if (file) foldersMap.set('/', [...(foldersMap.get('/') || []), file]);
            }
        }

        await Promise.all(promises);
        // console.log('foldersMap:', foldersMap);
        this.renderFillingModal(foldersMap);
    }

    private async readDirectory(dirEntry: any, parentPath: string = '', folderMap = new Map<string, File[]>())
        : Promise<Map<string, File[]>> {
        const currentPath = parentPath ? `${parentPath}/${dirEntry.name}` : '/' + dirEntry.name;
        if (!folderMap.has(currentPath)) {
            folderMap.set(currentPath, []);
        }

        return new Promise((resolve, reject) => {
            const reader = dirEntry.createReader();

            const readAllEntries = () => {
                reader.readEntries(async (entries: any[]) => {
                    if (entries.length === 0) {
                        resolve(folderMap);
                        return;
                    }
                    for (const entry of entries) {
                        if (entry.isFile) {
                            await new Promise<void>((done) => {
                                entry.file((file: File) => {
                                    folderMap.get(currentPath)!.push(file);
                                    done();
                                }, reject);
                            });
                        } else if (entry.isDirectory) {
                            await this.readDirectory(entry, currentPath, folderMap);
                        }
                    }
                    readAllEntries(); // keep reading until empty
                }, reject);
            };
            readAllEntries();
        });
    }
}