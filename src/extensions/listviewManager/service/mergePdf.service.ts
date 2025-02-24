import PAService from "./authService.service";
import { getSP, getSPByPath } from "../../../pnpjs-config";
import Swal from "sweetalert2";
import { SPFI } from "@pnp/sp";
import { ListViewCommandSetContext } from "@microsoft/sp-listview-extensibility";

export async function ConvertToPdf(context: ListViewCommandSetContext, selectedItems: any): Promise<Blob> {
    const paService = new PAService(
        context,
        'https://prod-236.westeurope.logic.azure.com:443/workflows/299c799fbb3247548c59280533f7506f/triggers/manual/paths/invoke?api-version=2016-06-01'
    );
    const sp = getSP(context);
    const spPortal = getSPByPath("https://epstin100.sharepoint.com/sites/EpsteinPortal", context);

    const requestBody: any[] = [];
    const missingFiles: any[] = [];

    for (const item of selectedItems) {
        const serverRelativeUrl = `https://epstin100.sharepoint.com${item.RelativePath}`;
        const exists = await checkIfFileOrFolderExists(item.RelativePath, sp);

        if (!exists) {
            missingFiles.push(item.Title);
            continue; // Skip this item since it doesn't exist
        }

        const siteAddress: string = await sp.site.getWebUrlFromPageUrl(serverRelativeUrl);

        const fullRelativePath: string = siteAddress.replace("https://epstin100.sharepoint.com", "");
        const relativePath: string = item.RelativePath.replace(fullRelativePath, "");

        requestBody.push({
            siteAddress,
            title: item.Title,
            relativePath
        });
    }

    // If there are missing files, show the confirmation modal
    if (missingFiles.length > 0) {
        const { isConfirmed } = await Swal.fire({
            title: "קבצים חסרים",
            html: `הקבצים הבאים לא קיימים במערכת ולכן לא ייכללו ב-PDF:<br><br> <b>${missingFiles.join("<br>")}</b>`,
            icon: "warning",
            showCancelButton: true,
            confirmButtonText: "כן, צור PDF",
            cancelButtonText: "לא, בטל",
            reverseButtons: true,
            didOpen: () => {
                const popup = document.querySelector(".swal2-container") as HTMLElement;
                if (popup) {
                    popup.style.zIndex = "1300"; // Set zIndex without !important
                }
            }
        });

        if (!isConfirmed) {
            throw new Error("המשתמש ביטל את יצירת ה-PDF בגלל קבצים חסרים.");
        }
    }

    const response: any = await paService.post(paService.CONVERT_TO_PDF, requestBody);

    try {
        const fileId = response.data.id;

        // If you want to download the file first, here's how:
        const file = await spPortal.web.getFileByServerRelativePath(decodeURIComponent(`/sites/EpsteinPortal${response.data.path}`)).getBuffer();
        const pdfBlob = new Blob([file], { type: 'application/pdf' });

        // Delete the file using getById and file ID
        await spPortal.web.lists.getById("1dd79a87-229c-4eaa-b15e-9b0afa29e56d").items.getById(fileId).delete();

        return pdfBlob;

    } catch (error) {
        throw new Error("שגיאה בהסרת הקובץ מה-SharePoint");
    }
}

export function downloadPdf(pdfBlob: Blob, fileName: string = "convertedFile.pdf") {
    const pdfUrl = URL.createObjectURL(pdfBlob);

    const a = document.createElement("a");
    a.href = pdfUrl;
    a.download = fileName;
    a.style.display = "none";
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(pdfUrl);
}

export async function checkIfFileOrFolderExists(serverRelativeUrl: string, sp: SPFI): Promise<boolean> {
    try {
        await sp.web.getFileByServerRelativePath(serverRelativeUrl)();
        return true;
    } catch (error: any) {
        console.error("Error checking file/folder existence:", error);
        return false;
    }
}



