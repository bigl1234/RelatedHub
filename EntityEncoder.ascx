<%@ Control Language="C#" AutoEventWireup="true" CodeFile="EntityCoder.ascx.cs" Inherits="RockWeb.Blocks.EntitySharing.EntityCoder" %>

<asp:UpdatePanel ID="upnlContent" runat="server">
    <ContentTemplate>

        <asp:Panel ID="pnlView" runat="server" CssClass="panel panel-block">

            <div class="panel-heading">
                <h1 class="panel-title"><i class="fa fa-share-square"></i>Rock Entity Export and Import</h1>
            </div>
            <div class="panel-body">                
                <Rock:NotificationBox ID="nbInfo" ClientIDMode="Static" runat="server" NotificationBoxType="Info" />                                

                <div class="row">
                    <div class="col-md-6">
                        <div class="form-group">
                            <h5>Entity</h5>
                            <asp:Literal ID="lEntityPickerLabel" runat="server">
                            <h5>Select Entity Type</h5>
                            </asp:Literal>
                            <Rock:RockDropDownList ID="ddlEntityTypePicker" runat="server" DataTextField="FriendlyName" DataValueField="Id" SelectionMode="Single" AppendDataBoundItems="true"
                                DisplayEnhancedAsAbsolute="true" EnhanceForLongLists="true" AutoPostBack="true" OnSelectedIndexChanged="ddlEntityTypePicker_SelectedIndexChanged">
                                <asp:ListItem Text="" Value="0"></asp:ListItem>
                            </Rock:RockDropDownList>

                            <Rock:RockListBox ID="rlbExportEntity" runat="server" Visible="true" DataTextField="Name" DataValueField="Id" AutoPostBack="false">
                            </Rock:RockListBox>

                            <Rock:Grid ID="gEntityTypePropertyList" runat="server" AutoGenerateColumns="false" AllowPaging="false" OnRowDataBound="gEntityTypePropertyList_RowDataBound">
                                <Columns>
                                    <Rock:RockTemplateField>
                                        <ItemTemplate>
                                            <asp:Literal ID="lPropertyTypeIcons" runat="server"></asp:Literal>
                                        </ItemTemplate>
                                    </Rock:RockTemplateField>
                                    <Rock:RockTemplateField>
                                        <ItemTemplate>
                                            <Rock:BootstrapButton ID="btnAddPropertyPath" runat="server" CssClass="btn btn-xs btn-primary" OnClick="btnAddPropertyPath_Click">
                                                <i class="fa fa-plus"></i>
                                            </Rock:BootstrapButton>
                                            <Rock:BootstrapButton ID="btnRemovePropertyPath" runat="server" CssClass="btn btn-xs btn-primary" OnClick="btnRemovePropertyPath_Click">
                                            <i class="fa fa-minus"></i>
                                            </Rock:BootstrapButton>
                                        </ItemTemplate>
                                    </Rock:RockTemplateField>

                                    <Rock:RockTemplateField>
                                        <ItemTemplate>
                                            <Rock:BootstrapButton ID="btnRockEntityInfo" runat="server" OnClick="btnRockEntityInfo_Click" CssClass="btn btn-xs btn-default" ToolTip="">
                                                                <i class="fa fa-code"></i>
                                            </Rock:BootstrapButton>
                                        </ItemTemplate>
                                    </Rock:RockTemplateField>

                                    <Rock:RockBoundField DataField="Name" HeaderText="Name"></Rock:RockBoundField>
                                    <Rock:RockBoundField DataField="PropertyType" HeaderText="PropertyType"></Rock:RockBoundField>

                                </Columns>
                            </Rock:Grid>

                        </div>

                    </div>
                    <div class="col-md-6">
                        <h5>Entity Paths</h5>

                        <Rock:CodeEditor ID="ceEntityMapTesting" runat="server" EditorMode="Lava" EditorTheme="VibrantInkDark" EditorHeight="800">

                        </Rock:CodeEditor>
                    </div>
                </div>

                <Rock:ModalDialog ID="mdlRelatedEntityInfo" runat="server" Title="Related Entity Property Information">
                    <Content>
                        <div class="row">
                            <div class="col-md-4">
                                <asp:Literal ID="lRelatedEntityTypeName" runat="server"></asp:Literal>
                            </div>
                            <div class="col-md-4 text-center">
                            </div>
                            <div class="col-md-4 text-right">
                                <%# DateTime.Now.ToShortTimeString() %>
                            </div>
                        </div>
                        <div class="row">
                            <div class="col-md-12">
                                Current Path: <asp:Literal ID="lModalBroswingPath" runat="server"></asp:Literal>
                                <br />
                                Coder Path: <asp:Literal ID="lModalCoderBroswingPath" runat="server"></asp:Literal>
                                                    
                            </div>
                        </div>
                        <div class="row">
                            <div class="col-md-12">
                                <Rock:NotificationBox ID="nbForbiddenRelatedEntity" ClientIDMode="Static" runat="server" NotificationBoxType="Warning" />
                            </div>
                        </div>
                        <div class="row">
                            <div class="col-md-12">
                                <Rock:Grid ID="gRelatedEntityTypeDetails" runat="server" AutoGenerateColumns="false" AllowPaging="false">
                                    <Columns>
                                        <Rock:SelectField DataVisibleField="Enabled"></Rock:SelectField>

                                        <Rock:RockTemplateField>
                                            <ItemTemplate>
                                                <Rock:BootstrapButton ID="btnRockEntityInfo" runat="server" OnClick="btnRockEntityInfo_Click" CssClass="btn btn-xs btn-default" ToolTip="">
                                                                    <i class="fa fa-code"></i>
                                                </Rock:BootstrapButton>
                                            </ItemTemplate>
                                        </Rock:RockTemplateField>

                                        <Rock:RockBoundField DataField="Name" HeaderText="Name"></Rock:RockBoundField>
                                        <Rock:RockBoundField DataField="TypeName" HeaderText="TypeName"></Rock:RockBoundField>
                                        <Rock:RockBoundField DataField="IsRockClass" HeaderText="IsRockClass"></Rock:RockBoundField>
                                        <Rock:RockBoundField DataField="ExportPropertyType" HeaderText="ExportPropertyType"></Rock:RockBoundField>

                                    </Columns>
                                </Rock:Grid>
                            </div>
                        </div>
                    </Content>
                </Rock:ModalDialog>
            </div>
        </asp:Panel>

        <asp:TreeView ID="tvSelectedEntities" runat="server" ShowCheckBoxes="Leaf" OnTreeNodeExpanded="tvSelectedEntities_TreeNodeExpanded">
        </asp:TreeView>
    </ContentTemplate>
</asp:UpdatePanel>



<script>

</script>
