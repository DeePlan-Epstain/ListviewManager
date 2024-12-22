import { ListViewCommandSetContext } from "@microsoft/sp-listview-extensibility";

export function navigateToFavorites(context: ListViewCommandSetContext) {
    const tenant = window.location.hostname.split('.')[0]; // Extract tenant dynamically
    const userName = context.pageContext.user.email.split('@')[0]; // Extract current user's username
    const url = `https://${tenant}-my.sharepoint.com/personal/${userName}/_layouts/15/onedrive.aspx?view=favorites`;
    window.location.href = url;
}