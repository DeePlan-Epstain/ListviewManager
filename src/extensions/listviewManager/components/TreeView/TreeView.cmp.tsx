import React from "react";
// import "./NavTreeStyle.scss";
import "./Loader.module.scss";
import { CacheProvider } from "@emotion/react";
import createCache from "@emotion/cache";
import { prefixer } from "stylis";
import rtlPlugin from "stylis-plugin-rtl";
import { ChevronLeft, Delete, UnfoldLess } from "@mui/icons-material";
import { RichTreeView, useTreeItem2, TreeItem2Provider, TreeItem2Root, TreeItem2Content, TreeItem2IconContainer, TreeItem2Label, TreeItem2DragAndDropOverlay } from "@mui/x-tree-view";
import { Collapse, IconButton, Link, TextField, Tooltip } from "@mui/material";
import { CustomTreeItem, LoadAsyncLibraries } from "../../models/TreeView.mdl";
import { NavTreeService } from "../../service/TreeView.srv";
import RefreshIcon from '@mui/icons-material/Refresh';

const cacheRtl = createCache({
    key: "muirtl",
    stylisPlugins: [prefixer, rtlPlugin],
});

export interface INavTreeProps {
    WebUri: string;
    context: any;
}

export interface INavTreeStates {
    TreeItems: CustomTreeItem[];
    AllItems: CustomTreeItem[];
    filteredTreeItems: CustomTreeItem[] | null;
    ExpandedNodes: string[];
    SkipExpandLibraries: string[];
    SelectedNode: string;
    filterValue: string;
    NodesServerRelativeUrls: Array<any>;
    LoadAsyncLibraries: LoadAsyncLibraries[];
    IsLoading: boolean;
    isRedirect: boolean;
    isFetching: boolean;
    isDisabled: boolean;
}

export class NavTree extends React.Component<INavTreeProps, INavTreeStates> {
    private navTreeService = new NavTreeService(this.props.context);

    constructor(props: INavTreeProps) {
        super(props);
        // Set States (information managed within the component), When state changes, the component responds by re-rendering
        this.state = {
            // Loader indicator.
            IsLoading: true,

            //is Redirect
            isRedirect: false,

            isFetching: false,

            isDisabled: false,

            // All tree links and hirerachy.
            TreeItems: [],

            AllItems: [],

            // Filtered tree items.
            filteredTreeItems: null,

            // Libraries to skip expanding(fetching their folders again).
            SkipExpandLibraries: [],

            // Libraries that we have to load async, after we are done fetching all the other libraries.
            LoadAsyncLibraries: [],

            // All the current node ids that are expanded.
            ExpandedNodes: [],

            // The currently selected node.
            SelectedNode: "",

            // filter value
            filterValue: "",

            // All the serverRelativeUrls of all the nodes that were already fetched.
            NodesServerRelativeUrls: [],
        };
        this.handleSiteClick = this.handleSiteClick.bind(this); // Binding method to `this`
    }

    componentDidMount() {
        this.init();
    }

    componentWillUnmount(): void {
        if (
            window.location.href === this.props.context.pageContext.web.absoluteUrl
        ) {
            window.sessionStorage.removeItem("ExpandedNodes");
        }
    }

    private init = async () => {
        let NewState = { ...this.state };
        // Get cached tree if exists in session storage.
        const cachedData = this.GetCachedItem("TreeItems");

        if (cachedData) {
            NewState = {
                ...this.state,
                IsLoading: false,
                TreeItems: cachedData.TreeItems,
                AllItems: cachedData.AllItems,
                NodesServerRelativeUrls: cachedData.NodesServerRelativeUrls,
                SkipExpandLibraries: cachedData.SkipExpandLibraries,
            };

            // Get expanded nodes if there are.
            if (sessionStorage.getItem("ExpandedNodes")) {
                const ExpandedAndSelected = this.GetCachedItem("ExpandedNodes");

                NewState = {
                    ...NewState,
                    ExpandedNodes: ExpandedAndSelected.ExpandedNodes,
                    SelectedNode: ExpandedAndSelected.SelectedNode,
                };
            }

            this.setState(NewState, () => {
                // Check if any library has to load async.
                if (
                    sessionStorage.getItem("LoadAsyncLibraries") &&
                    JSON.parse(sessionStorage.getItem("LoadAsyncLibraries") || "").length
                ) {
                    this.LoadLibrariesAsync();
                }
            });
        } else {
            // Fetch TreeItems.
            await this.getInitialState();
            this.SetCachedTreeItems("Avoid");
        }
    };

    private getInitialState = async () => {
        console.time("getInitialState");
        try {
            const TreeItems = await this.navTreeService.getInitialState();
            const AllItems = this.navTreeService.flattenTreeItems(TreeItems, this.state.AllItems);

            this.setState({ TreeItems, AllItems, IsLoading: false });
            console.timeEnd("getInitialState");
        } catch (err) {
            console.error("getInitialState Error:", err);
        }
    };

    private resetNavTree = () => {
        this.setState({
            isDisabled: true,
            IsLoading: true,
            isRedirect: false,
            isFetching: false,
            TreeItems: [],
            AllItems: [],
            filteredTreeItems: null,
            SkipExpandLibraries: [],
            LoadAsyncLibraries: [],
            ExpandedNodes: [],
            SelectedNode: "",
            filterValue: "",
            NodesServerRelativeUrls: []
        });
        sessionStorage.removeItem("TreeItems");
        this.init();
        setTimeout(() => {
            this.setState({ isDisabled: false }); // Re-enable the button after 30 seconds
        }, 30000); // 30 seconds
    };

    private LoadLibrariesAsync = async () => {
        try {
            const res = await this.navTreeService.loadLibrariesAsync(
                this.state.TreeItems,
                this.state.ExpandedNodes
            );

            if (!res) throw new Error("LoadLibrariesAsync Error");

            this.setState(
                {
                    TreeItems: res.TreeItems,
                    LoadAsyncLibraries: res.LoadAsyncLibraries,
                    ExpandedNodes: res.ExpandedNodes,
                },
                () => this.SetCachedTreeItems("Avoid")
            );
        } catch (error) { }
    };

    private GetCachedItem = (ItemName: string) => {
        const CacheToParse = window.sessionStorage.getItem(ItemName);

        if (!CacheToParse) return;

        return JSON.parse(CacheToParse);
    };

    private SetCachedTreeItems = async (Link: string) => {
        // This function maps all required states and sets them to sessionStorage.
        const TreeCache = {
            TreeItems: this.state.TreeItems,
            AllItems: this.state.AllItems,
            NodesServerRelativeUrls: this.state.NodesServerRelativeUrls,
            SkipExpandLibraries: this.state.SkipExpandLibraries,
        };

        const ExpandedAndSelected = {
            SelectedNode: this.state.SelectedNode,
            ExpandedNodes: this.state.ExpandedNodes,
        };

        const TreeItemsToSet = JSON.stringify(TreeCache);
        const ExpandedNodesToSet = JSON.stringify(ExpandedAndSelected);
        const LoadAsyncLibrariesToSet = JSON.stringify(
            this.state.LoadAsyncLibraries
        );

        window.sessionStorage.setItem("TreeItems", TreeItemsToSet);
        window.sessionStorage.setItem("ExpandedNodes", ExpandedNodesToSet);
        window.sessionStorage.setItem("LoadAsyncLibraries", LoadAsyncLibrariesToSet);

        if (Link !== "Avoid") this.openLink(Link);

        return "Cache was set";
    };

    private openLink = (Link: string, isBlank?: boolean) => {
        window.open(Link, isBlank ? "_blank" : "_self");
    }

    private GetNodeRecursively = (Tree: any, NodeId: string) => {
        for (var i = 0; i < Tree.length; i++) {
            const CurrNode: CustomTreeItem = Tree[i];

            // If the wanted node is found.
            if (NodeId === CurrNode.id) {
                return CurrNode;
                // If wanted node wasn't found, continue looking in it's children array.
            } else if (CurrNode?.children?.length) {
                const FoundNode: any = this.GetNodeRecursively(
                    CurrNode.children,
                    NodeId
                );
                if (FoundNode) return FoundNode;
            }
        }
    };

    // events

    private handleNode = async (TreeItemId: string) => {
        const { TreeItems, NodesServerRelativeUrls, filterValue, AllItems, isRedirect, ExpandedNodes, isFetching } = this.state;

        if (isFetching) return;

        let TreeItemsCopy = JSON.parse(JSON.stringify(TreeItems));
        let Folders;

        const folder: CustomTreeItem | undefined = AllItems.find((item: any) => item.id === TreeItemId);

        const UpdatedExpandedNodes = ExpandedNodes.includes(TreeItemId) ? ExpandedNodes.filter(n => n !== TreeItemId) : [...ExpandedNodes, TreeItemId];

        this.setState({ ExpandedNodes: UpdatedExpandedNodes, isFetching: true }, async () => {
            if (!folder || folder.type === "site" || folder.type === "subsite") {
                this.setState({ isFetching: false });
                return;
            } else Folders = await this.navTreeService.GetFolderChildrenFolders(folder?.link);

            // Replace the new folders with old.
            const UpdatedTreeItems: any = this.navTreeService.GetUpdatedTreeItemsRecursively(TreeItemsCopy, TreeItemId, Folders);
            const newTreeItems = [...this.state.AllItems, ...Folders];

            // Add expanded node to fetched nodes array.
            const UpdatedNodesServerRelativeUrl = this.navTreeService.GetUpdatedSavedNodesServerRelativeUrl(TreeItemId, folder.link, NodesServerRelativeUrls);


            this.setState({ TreeItems: UpdatedTreeItems, AllItems: newTreeItems, NodesServerRelativeUrls: UpdatedNodesServerRelativeUrl, isFetching: false }, async () => {
                isRedirect ? this.SetCachedTreeItems(folder.link) : this.SetCachedTreeItems("Avoid");
                if (this.state.filteredTreeItems?.length) this.filterTreeItems(filterValue, true);
            });
        });
    };

    private HandleOnLabel = (ItemId: string) => {
        this.setState({ isRedirect: true }, () => this.handleNode(ItemId));
    }

    private handleSiteClick = (Link: string) => {
        this.openLink(Link);
        this.SetCachedTreeItems("Avoid");
    };

    private onNodeSelect = (_: any, SelectedNode: any) => {
        this.setState({ SelectedNode }, () => this.SetCachedTreeItems("Avoid"));
    };

    private filterTreeItems = (searchValue: string, keepExpanded?: boolean): void => {
        const { TreeItems, ExpandedNodes } = this.state;

        const expandedNodes = new Set<string>();

        if (!searchValue)
            return this.setState({ filteredTreeItems: null, ExpandedNodes: [] });

        const filteredTreeItems = this.navTreeService.filterTreeItems(
            TreeItems,
            searchValue,
            expandedNodes
        );

        this.setState({
            filteredTreeItems,
            ExpandedNodes: keepExpanded ? ExpandedNodes : Array.from(expandedNodes),
            filterValue: searchValue,
        });
    };

    private onCloseExpanded = () => {
        this.setState({ ExpandedNodes: [], SelectedNode: "" }, () => {
            this.SetCachedTreeItems("Avoid");
        });
    };

    // private collapseNode = (itemId: string) => {
    //   const isExpanded = this.state.ExpandedNodes.includes(itemId);

    //   if (!isExpanded) return;

    //   const { ExpandedNodes } = this.state;
    //   const UpdatedExpandedNodes = ExpandedNodes.filter((n) => n !== itemId);

    //   this.setState({ ExpandedNodes: UpdatedExpandedNodes }, () => {
    //     this.SetCachedTreeItems("Avoid");
    //   });
    // };

    private collapseNode = (itemId: string) => {
        const { ExpandedNodes } = this.state;

        // Find the index of the itemId
        const index = ExpandedNodes.indexOf(itemId);

        // If itemId is not found, do nothing
        if (index === -1) return;

        // Remove itemId and all subsequent IDs
        const UpdatedExpandedNodes = ExpandedNodes.slice(0, index);

        this.setState({ ExpandedNodes: UpdatedExpandedNodes }, () => {
            this.SetCachedTreeItems("Avoid");
        });
    };

    private CustomTreeItem = React.forwardRef((props: CustomTreeItem, ref: React.Ref<HTMLLIElement>): JSX.Element => {
        const { id, itemId, label, children, ...other } = props;

        const {
            getRootProps,
            getContentProps,
            getIconContainerProps,
            getLabelProps,
            getGroupTransitionProps,
            getDragAndDropOverlayProps,
            status,
            publicAPI,
        } = useTreeItem2({ id, itemId, children, label, rootRef: ref });

        const item: CustomTreeItem = (publicAPI as any).getItem(itemId);
        const folderType = item?.type;

        // Icon based on the folder type (for example 'folder' type might use a custom icon)
        const folderIcon = folderType ? (
            <span className={`TreeItemIcon TreeItemIcon_${folderType}`}></span>
        ) : null;

        // Ensure both icons are passed to the labelProps
        const labelProps = getLabelProps({
            icon: <>{folderIcon}</>,
            expandable: status.expanded.toString()
        });

        const isSite = item.type === 'site' || item.type === 'subsite' || item.type === 'library';
        const onLabelClick = () => {
            this.setState({ isRedirect: true })
            if (isSite) {
                this.handleNode(itemId);
                this.handleSiteClick(item.link);
            } else this.HandleOnLabel(item.id);

        };

        //  loading children
        if (item?.type === "loading") {
            return (
                <div className="TreeViewLink">
                    <div className="spinner">
                        <div className="bounce1"></div>
                        <div className="bounce2"></div>
                        <div className="bounce3"></div>
                    </div>
                </div>
            );
        }

        return (
            // @ts-ignore
            <TreeItem2Provider itemId={itemId}>
                <TreeItem2Root {...getRootProps(other)}>
                    <TreeItem2Content sx={{ padding: "0 8px" }} {...getContentProps()}>
                        <TreeItem2IconContainer {...getIconContainerProps()}
                            onClick={() => {
                                if (this.state.isFetching) return;
                                this.handleNode(itemId);
                                this.collapseNode(item.id);
                            }}>
                            {children && (
                                <ChevronLeft
                                    sx={{
                                        transform: status.expanded
                                            ? "rotate(-90deg)"
                                            : "rotate(0deg)",
                                        transition: "transform 0.3s ease-in-out",
                                        opacity: this.state.isFetching ? 0.5 : 1
                                    }}
                                />
                            )}
                        </TreeItem2IconContainer>
                        <TreeItem2Label
                            sx={{
                                display: "flex",
                                alignItems: "center",
                                fontSize: "14px",
                                color: "#666",
                                textAlign: "left",
                                padding: "4px 0"
                            }}
                            onClick={onLabelClick}
                            {...labelProps}
                        >
                            {folderIcon}
                            {label}
                        </TreeItem2Label>
                        <TreeItem2DragAndDropOverlay {...getDragAndDropOverlayProps()} />
                    </TreeItem2Content>
                    {children && (
                        <Collapse
                            {...getGroupTransitionProps()}
                            style={{ paddingInlineStart: "10px" }}
                        />
                    )}
                </TreeItem2Root>
            </TreeItem2Provider>
        );
    });

    public render(): React.ReactElement<INavTreeProps> {
        const { TreeItems, filteredTreeItems, ExpandedNodes, IsLoading, isFetching, isDisabled } = this.state;

        const Loader = () => (
            <div className="loading loading03">
                <span>ט</span>
                <span>ו</span>
                <span>ע</span>
                <span>ן</span>
                <span className="transparent">_</span>
                <span>ע</span>
                <span>ץ</span>
                <span className="transparent">_</span>
                <span>נ</span>
                <span>י</span>
                <span>ו</span>
                <span>ו</span>
                <span>ט</span>
                <span>.</span>
                <span>.</span>
                <span>.</span>
            </div>
        );

        return (
            <div dir="rtl" className="LeftNavTreeContainer" >
                <CacheProvider value={cacheRtl}>

                    {/* Actions */}
                    {!!TreeItems.length && (
                        <div className="TreeActionsContainer">
                            <TextField
                                label="חיפוש בעץ"
                                style={{ flex: 1 }}
                                onChange={(e) => this.filterTreeItems(e.target.value)}
                                variant="standard"
                            />

                            <Tooltip title={"רענן עץ"} placement="top">
                                <IconButton
                                    onClick={this.resetNavTree}
                                    disabled={isDisabled}
                                >
                                    <RefreshIcon />
                                </IconButton>
                            </Tooltip>

                            <Tooltip title="מזער" placement="top">
                                <IconButton
                                    disabled={!ExpandedNodes?.length || !!filteredTreeItems?.length || isFetching}
                                    onClick={this.onCloseExpanded}
                                >
                                    <UnfoldLess />
                                </IconButton>
                            </Tooltip>
                        </div>
                    )}

                    {IsLoading ? (
                        <Loader />
                    ) : (
                        <RichTreeView
                            expandedItems={ExpandedNodes}
                            //onItemExpansionToggle={(_, itemId, isExpanded) => isExpanded ? this.handleNode(itemId) : null}
                            onSelectedItemsChange={this.onNodeSelect}
                            items={filteredTreeItems || TreeItems}
                            slots={{ item: this.CustomTreeItem }}
                            classes={{ root: "TreeNavView" }}
                        />
                    )}

                    {/* Link to recycle bin */}
                    {!!TreeItems.length && (
                        <Link
                            className="RecycleBinLink"
                            href={`${this.props.WebUri}/_layouts/15/RecycleBin.aspx`}
                            data-interception="off"
                            style={{ padding: "0 20px 0 0", margin: "0 20px 1em 0" }}
                            onClick={() => this.openLink(`${this.props.WebUri}/_layouts/15/RecycleBin.aspx`)}>
                            <Delete />
                            סל מחזור
                        </Link>
                    )}
                </CacheProvider>
            </div>
        );
    }
}
