import * as React from 'react';
import { IGlobalNavProps } from "./IGlobalNavProps";
import { IGobalNavState } from "./IGlobalNavState";
import { sp } from "@pnp/sp";
import "@pnp/sp/taxonomy";
import { dateAdd, PnPClientStorage } from "@pnp/common";
import { IOrderedTermInfo } from '@pnp/sp/taxonomy';
import { ContextualMenuItemType, IButtonStyles } from '@fluentui/react';
import { createTheme, ITheme } from 'office-ui-fabric-react/lib/Styling';
import { CommandBar, ICommandBarStyleProps } from "@fluentui/react/lib/CommandBar";

const myKey: string = "navigationElements";

const theme: ITheme = createTheme({
    semanticColors: {
        bodyBackground: "#333",
        bodyText: "#fff"
    }
});
const CommandBarProps: ICommandBarStyleProps = {
    theme: theme
};
const buttonStyle: IButtonStyles = {
    root: {
        backgroundColor: "#333",
        color: "#fff"
    },
    menuIcon: {
        color: "#fff",
    }
};

export default class GlobalNav extends React.Component<IGlobalNavProps, IGobalNavState> {
    private store = new PnPClientStorage();
    constructor(props: IGlobalNavProps) {
        super(props);
        this.state = {
            loading: false,
            terms: []
        }
    }
    public componentDidMount() {
        this.setState({}, async () => {
            // this portion is responsible for getting terms from term store
            const cachedTermInfo = await this.store.local.getOrPut(myKey, () => {
                return sp.termStore.groups.getById(this.props.termGroupId).sets.getById(this.props.termSetId).getAllChildrenAsOrderedTree({ retrieveProperties: true });
            }, dateAdd(new Date(), "minute", 1));
            if (cachedTermInfo.length > 0) {
                console.log(cachedTermInfo);
                this.setState({ terms: cachedTermInfo });
            }
        });
    }

    private menuItems(menuItem: any, itemType: ContextualMenuItemType) {
        return ({
            key: menuItem.id,
            name: menuItem.defaultLabel,
            itemType: itemType,
            href: menuItem.children.length == 0 ?
                ((menuItem.localProperties != undefined && menuItem.localProperties[0].properties !== undefined && menuItem.localProperties[0].properties.length > 0) ?
                    menuItem.localProperties[0].properties.filter(x => x.key == "_Sys_Nav_SimpleLinkUrl")[0].value !== undefined ? menuItem.localProperties[0].properties.filter(x => x.key == "_Sys_Nav_SimpleLinkUrl")[0].value : null
                    : null)
                : null,
            subMenuProps: menuItem.children.length > 0 ?
                { items: menuItem.children.map((i) => { return (this.menuItems(i, ContextualMenuItemType.Normal)); }) }
                : null,
            isSubMenu: itemType != ContextualMenuItemType.Header,
            buttonStyles: buttonStyle
        });
    }

    public render(): React.ReactElement<IGlobalNavProps> {
        var commandBarItems: any[] = [];
        if (this.state.terms.length > 0) {
            commandBarItems = this.state.terms.map((i) => {
                return (this.menuItems(i, ContextualMenuItemType.Header));
            });
        }
        return (
            <>
                {
                    this.state.terms.length > 0 &&
                    <div>
                        <CommandBar  {...CommandBarProps}
                            style={{ width: "100%" }}
                            items={commandBarItems}
                        />
                    </div>
                }
            </>
        );
    }
    
}