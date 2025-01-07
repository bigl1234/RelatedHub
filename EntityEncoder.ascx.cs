// <copyright>
// Copyright by the Spark Development Network
//
// Licensed under the Rock Community License (the "License");
// you may not use this file except in compliance with the License.
// You may obtain a copy of the License at
//
// http://www.rockrms.com/license
//
// Unless required by applicable law or agreed to in writing, software
// distributed under the License is distributed on an "AS IS" BASIS,
// WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
// See the License for the specific language governing permissions and
// limitations under the License.
// </copyright>
//
using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using System.Data.Entity.Core.Mapping;
using System.Data.SqlClient;
using System.Linq;
using System.Management.Automation.Runspaces;
using System.Reflection;
using System.Runtime.Serialization;
using System.Text;
using System.Web.UI;
using System.Web.UI.HtmlControls;
using System.Web.UI.WebControls;
using AngleSharp.Text;
using DocumentFormat.OpenXml.Wordprocessing;
using ImageResizer;
using Lucene.Net.Support;
using Mono.CSharp;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Text;
using OpenXmlPowerTools;
using Quartz.Util;
using Rock;
using Rock.Address;
using Rock.Attribute;
using Rock.Data;
using Rock.Lava;
using Rock.Model;
using Rock.Utility.EntityCoding;
using Rock.ViewModels.Utility;
using Rock.Web;
using Rock.Web.Cache;
using Rock.Web.UI;
using Rock.Web.UI.Controls;
using Rock.Web.Utilities;

namespace RockWeb.Blocks.EntitySharing
{
    [BooleanField(
     "Always Export Required Entities",
     Description = "If set to true, all related entities required to insert the entity into the database will be exported?",
     DefaultBooleanValue = false,
     Key = AttributeKey.AlwaysExportRequiredEntities,
     Order = 1)]

    [KeyValueListField("Hidden Properties List",
        "A list of standard properties to hide from the results. Key is the property name and value is optional for the entity type.", true,
        "Id|CreatedByPersonAliasId|ModifiedByPersonAliasId|CreatedDateTime|ModifiedDateTime",
        Key = AttributeKey.HidePropertyList,
        Order = 4)]

    [KeyValueListField("Default Exclude Property List",
        "A list of standard properties to default to unchecked from the results. Key is the property name and value is optional for the entity type.", true,
        DefaultValue = "Id|ForeignKey|ForeignGuid|ForeignId|CreatedByPersonAliasId|ModifiedByPersonAliasId|CreatedDateTime|ModifiedDateTime",
        Key = AttributeKey.UnselectPropertyList,
        Order = 4)]

    [DisplayName("Entity Coding")]
    [Category("Entity Coder")]
    [Description("Import and export entities into and out of Rock.")]
    [Rock.SystemGuid.BlockTypeGuid("FED5A1A4-3A1E-4FF0-BDF2-F9542E85CE4E")]
    public partial class EntityCoder : Rock.Web.UI.RockBlock
    {
        #region Attribute Keys

        /// <summary>
        /// Keys to use for Block Attributes
        /// </summary>
        private static class AttributeKey
        {
            public const string AlwaysExportRequiredEntities = "AlwaysExportRequiredEntities";
            public const string HidePropertyList = "HidePropertyList";
            public const string UnselectPropertyList = "UnselectPropertyList";

        }

        protected string[] HidePropertyList
        {
            get
            {
                return GetAttributeValue(AttributeKey.HidePropertyList).Split('|');
            }
        }
        protected string[] UnselectPropertyList
        {
            get
            {
                return GetAttributeValue(AttributeKey.UnselectPropertyList).Split('|');
            }
        }

        #endregion

        #region View States

        /// <summary>
        /// Viewstate to store the properties for the entity type being browsed
        /// </summary>
        public List<PropertyPathModel> EntityProperties
        {
            get
            {
                var props = ViewState["EntityProperties"] as List<PropertyPathModel>;
                if (props != null)
                {
                    return props;
                }
                else
                {
                    ViewState["EntityProperties"] = new List<PropertyPathModel>();
                    return EntityProperties;
                }
            }
            set
            {
                if (value == null)
                {
                    var props = new List<PropertyPathModel>();
                    ViewState["EntityProperties"] = props;
                }
                else
                {
                    ViewState["EntityProperties"] = value;
                }
            }
        }

        /// <summary>
        /// Viewstate to store the entity type name for the entity type being browsed
        /// </summary>
        public string BrowsingEntityTypeName
        {
            get
            {

                if (ViewState["BrowsingEntityTypeName"] != null)
                {
                    return ViewState["BrowsingEntityTypeName"].ToString().Trim();
                }
                else
                {
                    return "";
                }
            }
            set
            {
                ViewState["BrowsingEntityTypeName"] = value;
            }
        }

        /// <summary>
        /// Viewstate to store the path to the entity type for the entity type being browsed
        /// </summary>
        public string BrowsingPath
        {
            get
            {

                if (ViewState["BrowsingPath"] != null)
                {
                    return ViewState["BrowsingPath"].ToString().Trim();
                }
                else
                {
                    return "";
                }
            }
            set
            {
                lModalBroswingPath.Text = value;
                ViewState["BrowsingPath"] = value;
            }
        }

        /// <summary>
        /// Viewstate to store the path the core EntityCoder will use to the entity type for the entity type being browsed
        /// </summary>
        public string CoderBrowsingPath
        {
            get
            {

                if (ViewState["CoderBrowsingPath"] != null)
                {
                    return ViewState["CoderBrowsingPath"].ToString().Trim();
                }
                else
                {
                    return "";
                }
            }
            set
            {
                lModalCoderBroswingPath.Text = value;
                ViewState["CoderBrowsingPath"] = value;
            }
        }

        /// <summary>
        /// A list of the <see cref="EntityPathModel"/> models of the user selected entity types to process.
        /// Always must have the root entity to be valid.
        /// </summary>
        public List<EntityPathModel> SelectedEntityPaths
        {
            get
            {

                if (ViewState["SelectedEntityPaths"] != null)
                {
                    return ViewState["SelectedEntityPaths"] as List<EntityPathModel>;
                }
                else
                {                    
                    return new List<EntityPathModel> {
                        new EntityPathModel{
                            CoderPath = "",
                            TypeName = BrowsingEntityTypeName,
                            TypePath = BrowsingEntityTypeName
                        }
                    };
                }
            }
            set
            {
                if (value == null)
                {
                    ViewState["SelectedEntityPaths"] = new List<string> { "" };
                }
                else
                {
                    ViewState["SelectedEntityPaths"] = value;
                }
            }
        }

        #endregion

        #region View Models

        [Serializable]
        public class PropertyPathModel
        {
       
            public string EntityTypeCoderPath { get; set; }

            /// <summary>
            /// Not the path that will be sent to the coder.
            /// This is just a path we use in this UI for selection purposes.
            /// </summary>
            public string EntityTypeName { get; set; }

            /// <summary>
            /// The name of the system property.
            /// This is usually a database table name but child entities will have the linkage as an Id on the child table.
            /// This process does use reflection so there may be some properties that are not true Rock database table properties,
            /// we try to filter these out and could 100% if we hit the database but has not been an issue in testing yet.
            /// </summary>
            public string Name { get; set; }

            /// <summary>
            /// This is needed because we are combining some properties into one property before displaying to the user.
            /// For instance, Page has a ParentPage property but it also has a ParentPageId property. In these instances the
            /// two properties are combined because it is really just one object. The core Rock coder will use the database
            /// table name as the path to the related entity rather than the virtual property name.
            /// </summary>
            public string CoderName { get; set; }

            /// <summary>
            /// The Rock Entity type name.
            /// Example Rock.Model.Page
            /// </summary>
            public string TypeName { get; set; }

            /// <summary>
            /// Is this property defined a value that is a Rock Entity type?
            /// </summary>
            public bool IsRockClass { get; set; }

            /// <summary>
            /// Is this property defining a value that is a colleciton of values.
            /// </summary>
            public bool IsArray { get; set; }

            /// <summary>
            /// A summarized version of the type of property. This was a bit difficult to get to so there may be some bugs here.
            /// "Rock Entity" - A related Rock entity. Usually found with an Id field on the entity. (ParentPageId)
            /// "Rock Entity Array" - An array of child Rock entities. (Page.Pages)
            /// "Property" - Database property 
            /// </summary>
            public string PropertyType { get; set; }

            /// <summary>
            /// Has the user selected this property to be included.?
            /// Defaulted based on general rules when selecting an entity type.
            /// </summary>
            public bool Selected { get; set; }

            /// <summary>
            /// Are we allowed to capture this data outside of Rock?
            /// Person, Finance, etc tables are excluded. We need to be careful of PII.
            /// We also need to be careful to not crash the system. Exporting the entire
            /// AttributeValues table would probably melt the server.
            /// </summary>
            public bool Enabled { get; set; }

            /// <summary>
            /// Is the entity type for this property sensitive or forbidden to export?
            /// </summary>
            public bool Forbidden { get; set; }
        }

        [Serializable]
        public class EntityPathModel
        {
            public string CoderPath { get; set; }
            public string ParentCoderPath { get; set; }
            public string TypePath { get; set; }
            public string TypeName { get; set; }
            public List<PropertyPathModel> Properties {get;set;}
            public List<EntityPathModel> RelatedEntityPaths { get; set; }
            public List<EntityPathModel> ChildEntityPaths { get; set; }

        }

        #endregion

        #region Properties


        public EntityPathModel BaseMap
        {
            get
            {
                if (SelectedEntityPaths.Where(m => m.CoderPath == "").Any())
                {
                    return SelectedEntityPaths.Where(m => m.CoderPath == "").First();
                }
                else if(!String.IsNullOrEmpty(BrowsingEntityTypeName))
                {
                    var map = new EntityPathModel();
                    map.CoderPath = "";
                    map.ChildEntityPaths = new List<EntityPathModel>();
                    map.Properties = new List<PropertyPathModel>();
                    map.RelatedEntityPaths = new List<EntityPathModel>();
                    map.TypePath = BrowsingEntityTypeName;
                    var entitytype = EntityTypeCache.Get(BrowsingEntityTypeName);
                    var props = EntityExport.GetEntityTypeProperties(entitytype, UnselectPropertyList, HidePropertyList);
                    map.Properties = ConvertEntityPropertiesToViewModel(props);
                    return map;

                }
                return null;
            }
        }


        #endregion

        #region Helper Classes

        /// <summary>
        /// 
        /// Building a usable file for the core Rock encoding peocess is a goal
        /// of this solution. Sharing entities can become a very powerful tool in
        /// the Rock developer's toolbox.
        /// 
        /// The string path to use to find the parent record.
        /// This is the core Path string where the base entity type path is an empty string.
        /// A "." pattern is used to indicate which property the related entities can be found.
        ///
        /// CHILD ENTITIES
        /// Example: Child Entities - We will use Rock.Model.Page as an example
        /// Page has a property of child page objects, the property name is "Pages".
        /// If Page was our parent and this property path is the Pages property our path to the parent entity would be an enpty string.
        /// Our path to the child pages is ".Pages" since the base is an empty string.
        ///
        /// If we were doing the same but say coming from Layout where Page is a child of layout the would be ".Pages.Pages" to the
        /// Page Pages property. Empty string indicating the root entity Layout, Layout has a property called Pages to find the child
        /// pages. Finally, Page has a property called Pages for it's child pages.
        ///
        /// We need to keep this path structure intact because the core Rock entity encoding processing uses it to determine
        /// which entities to restore and the relationships to set ids for new records being created.
        ///
        /// RELATED ENITIES
        /// For related entities it is the same but now we have a database property linking us directly to a single entity.
        /// This is much easier to deal with programatically.
        /// Example: To get to the Site from a Page the path would be ".LayoutId.SiteId". The core encoder tries to resolve
        /// int fields to a Rock entity type if it can. There may be some bugs here because the code looks like it was
        /// originally coded to match by Id being at then end of an int field and removing the Id matches a table name.
        /// But it seems to be working correctly and there is some reflection going on in the code base. Our implementation will
        /// check array entity types and property types from reflection to try and detect Rock entities.
        /// 
        /// </summary>



        #endregion

        #region Base Method Overrides

        /// <summary>
        /// Initialize basic information about the page structure and setup the default content.
        /// </summary>
        /// <param name="sender">Object that is generating this event.</param>
        /// <param name="e">Arguments that describe this event.</param>
        protected void Page_Load(object sender, EventArgs e)
        {
            //ScriptManager.GetCurrent( this.Page ).RegisterPostBackControl( btnExport );
            if (!Page.IsPostBack)
            {
                LiteralControl jsResource = new LiteralControl();
                jsResource.Text = "<script type=\"text/javascript\" src=\"https://ajax.googleapis.com/ajax/libs/jquery/2.1.1/jquery.min.js\"></script>";
                jsResource.Text += "<script type=\"text/javascript\" src=\"https://rawgit.com/abodelot/jquery.json-viewer/master/json-viewer/jquery.json-viewer.js\"></script>";
                Page.Header.Controls.Add(jsResource);

                HtmlLink stylesLink = new HtmlLink();
                stylesLink.Attributes["rel"] = "stylesheet";
                stylesLink.Attributes["type"] = "text/css";
                stylesLink.Href = "https://rawgit.com/abodelot/jquery.json-viewer/master/json-viewer/jquery.json-viewer.css";
                Page.Header.Controls.Add(stylesLink);

                ddlEntityTypePicker.DataSource =
                    EntityTypeCache.All()
                    .Where(et => EntityExport.ExportableEntityTypes.Contains(et.Name))
                    .OrderBy(et => et.Name);

                ddlEntityTypePicker.DataBind();
            }
        }

        #endregion

        #region Control Binding Methods

        /// <summary>
        /// Binds the preview grid.
        /// </summary>
        protected void BindSelectedPathsToTreeView(TreeNode callingnode = null)
        {

            if (callingnode == null)
            {
                tvSelectedEntities.Nodes.Clear();
                var basenode = new TreeNode();
                basenode.Text = BrowsingEntityTypeName;
                basenode.Value = EntityTypeCache.Get(BrowsingEntityTypeName).Name;
                basenode.Checked = true;
                basenode.Expanded = true;
                basenode.ShowCheckBox = false;
                basenode.NavigateUrl = "";

                var propertiesnode = new TreeNode();
                propertiesnode.Text = "Properties";
                propertiesnode.NavigateUrl = "Properties";
                propertiesnode.Expanded = true;
                basenode.ChildNodes.Add(propertiesnode);

                var rockentitynode = new TreeNode();
                rockentitynode.ShowCheckBox = false;                
                rockentitynode.Text = "Related Entities";
                rockentitynode.NavigateUrl = "Related Entities";
                rockentitynode.Expanded = true;
                rockentitynode.Value = BrowsingEntityTypeName;
                basenode.ChildNodes.Add(rockentitynode);

                var rockentitylistnode = new TreeNode();
                rockentitylistnode.ShowCheckBox = false;
                rockentitylistnode.Text = "Child Entities";
                rockentitylistnode.NavigateUrl = "Child Entities";
                rockentitylistnode.Expanded = true;
                rockentitylistnode.Value = BrowsingEntityTypeName;
                basenode.ChildNodes.Add(rockentitylistnode);

                var baseprops = EntityProperties.Where(p => p.EntityTypeCoderPath == CoderBrowsingPath && p.Name != "Id" && p.Name != "Guid");

                foreach (var prop in baseprops)
                {
                    var propnode = new TreeNode();
                    propnode.Text = prop.Name;
                    propnode.Value = prop.TypeName;
                    propnode.Checked = prop.Selected;

                    if (prop.IsRockClass == false)
                    {                        
                        propertiesnode.ChildNodes.Add(propnode);
                    }
                    else if (prop.IsArray == true)
                    {
                        var emptynode = new TreeNode();
                        emptynode.Value = "ChildEntitiesNode";
                        propnode.ChildNodes.Add(emptynode);
                        rockentitylistnode.ChildNodes.Add(propnode);
                    }
                    else
                    {
                        var emptynode = new TreeNode();
                        emptynode.Value = "RelatedentityNode";
                        propnode.ChildNodes.Add(new TreeNode());
                        rockentitynode.ChildNodes.Add(propnode);
                    }
                }


                tvSelectedEntities.Nodes.Add(basenode);
            }
            
            
        }

        #endregion

        #region Event Handlers

        protected EntityPathModel EntityMap(EntityPathModel pathModel = null)
        {
            var basemap = new EntityPathModel();

            if (pathModel == null)
            {
                basemap = SelectedEntityPaths.Where(m => m.CoderPath == "").First();
                basemap.RelatedEntityPaths = SelectedEntityPaths.Where(r => r.ParentCoderPath == basemap.CoderPath).ToList();
                foreach (var map in basemap.RelatedEntityPaths)
                {
                    EntityMap(map);
                }
            }
            else
            {
                foreach(var map in pathModel.RelatedEntityPaths)
                {
                    map.RelatedEntityPaths = SelectedEntityPaths.Where(r => r.ParentCoderPath == map.CoderPath).ToList();
                    EntityMap(map);
                }
            }

            return basemap;
        }

        protected void LoadEntityMapToCodeEditor(bool simplepropertylist = true)
        {
            ceEntityMapTesting.Text = EntityMap().ToJson(true);

            //if (simplepropertylist)
            //{

            //    var map = SelectedEntityPaths.Select(m => new {
            //        Path = m.CoderPath,
            //        m.TypePath,
            //        m.TypeName,
            //        Properties = EntityProperties.Where(p => p.EntityTypeCoderPath == m.CoderPath && p.EntityTypeName == m.TypeName)
            //            .Select(p => p.Name).ToArray()
            //    });

                
            //    ceEntityMapTesting.Text = map.ToJson(true);
            //}
            //else
            //{
            //    var entitymap = new
            //    {
            //        EntityPaths = SelectedEntityPaths.Select(t => new
            //        {
            //            Path = t.CoderPath,
            //            t.TypePath,
            //            t.TypeName,
            //            Properties = EntityProperties.Where(p => p.EntityTypeCoderPath == t.CoderPath).ToList()
            //        })
            //    };
            //    ceEntityMapTesting.Text = entitymap.ToJson(true);
            //}


            
        }

        protected void ddlEntityTypePicker_SelectedIndexChanged(object sender, EventArgs e)
        {
            nbInfo.Visible = false;
            BrowsingPath = "";
            BrowsingEntityTypeName = "";           

            var entitytypeid = ddlEntityTypePicker.SelectedValueAsId();
            if (entitytypeid > 0)
            {
                UpdateBaseEntityType();
                BindSelectedPathsToTreeView();
                LoadEntityMapToCodeEditor();
            }
            else
            {
                nbInfo.Text = "Entity Properties Not Loaded";
                nbInfo.NotificationBoxType = NotificationBoxType.Warning;
                nbInfo.Visible = true;
            }            
        }

        protected void btnRockEntityInfo_Click(object sender, EventArgs e)
        {

        }

        #endregion

        #region Custom Methods

        public List<PropertyPathModel> ConvertEntityPropertiesToViewModel(List<EntityProperty> properties)
        {            
            return properties.Select(p => new PropertyPathModel
                {
                    EntityTypeCoderPath = CoderBrowsingPath,
                    EntityTypeName = BrowsingPath,
                    Name = p.Name,
                    CoderName = p.DatabaseColumnName,
                    TypeName = p.ResolvedTypeFullName,
                    IsRockClass = p.IsRockClass,
                    IsArray = p.IsArray,
                    PropertyType = p.ExportPropertyType,
                    Selected = p.Selected,
                    Enabled = p.Enabled,
                    Forbidden = false
                        //p.ResolvedTypeFullName.StartsWith("Rock.Model.") &&
                        //ForbiddenEntityTypes.Where(f => f.Name == p.ResolvedTypeFullName).Any()

                }).ToList();
        }

        public void RefreshEntityTypeInformation(EntityTypeCache selectedentitytype)
        {            
            BrowsingEntityTypeName = selectedentitytype.Name;
            BrowsingPath = selectedentitytype.Name;
            CoderBrowsingPath = "";
            var props = EntityExport.GetEntityTypeProperties(selectedentitytype, UnselectPropertyList, HidePropertyList);
            var newprops = ConvertEntityPropertiesToViewModel(props);
            EntityProperties = newprops;
        }

        /// <summary>
        /// Loads the currently viewed entity type properties into view state
        /// </summary>
        /// <param name="entitytype"></param>
        private void UpdateBaseEntityType()
        {
            if (ddlEntityTypePicker.SelectedValueAsId().Value > 0)
            {
                SelectedEntityPaths = null;
                using (var context = new RockContext())
                {
                    var entityType = EntityTypeCache.Get(ddlEntityTypePicker.SelectedValueAsId().Value);
                    var serviceInstance = Reflection.GetServiceForEntityType(entityType.GetEntityType(), context);
                    var qryMethod = serviceInstance.GetType().GetMethod("Queryable", new Type[] { });
                    var idprop = entityType.GetEntityType().GetProperty("Id");
                    var nameprop = entityType.GetEntityType().GetProperty("Name");

                    var basemodel = BaseMap;
                    basemodel.Properties = EntityProperties.Where(p => p.EntityTypeCoderPath == CoderBrowsingPath).ToList();
                    SelectedEntityPaths = new List<EntityPathModel> { basemodel };

                    RefreshEntityTypeInformation(entityType);                                       

                    gEntityTypePropertyList.DataSource = basemodel.Properties;
                    gEntityTypePropertyList.DataBind();

                    nbInfo.Text = "Entity Properties Loaded";
                    nbInfo.NotificationBoxType = NotificationBoxType.Success;
                    nbInfo.Visible = true;
                }
            }
            else
            {
                nbInfo.Text = "Entity Properties Not Loaded";
                nbInfo.NotificationBoxType = NotificationBoxType.Warning;
                nbInfo.Visible = true;
            }
        }

        #endregion





        protected void btnAddPropertyPath_Click(object sender, EventArgs e)
        {
            nbInfo.Visible = false;
            var button = sender as BootstrapButton;            

            if (button != null && !String.IsNullOrEmpty(button.Attributes["PropertyName"]))
            {
                if (EntityProperties.Where(p => p.Name == button.Attributes["PropertyName"]).Any())
                {

                    var prop = EntityProperties.Where(p => p.Name == button.Attributes["PropertyName"]).First();
                    if (prop.TypeName.StartsWith("Rock.Model."))
                    {
                        var relatedentitytype = EntityTypeCache.Get(prop.TypeName);
                        var relatedentityprops = EntityExport.GetEntityTypeProperties(relatedentitytype, UnselectPropertyList, HidePropertyList);

                    }
                }
            }
        }

        protected void btnRemovePropertyPath_Click(object sender, EventArgs e)
        {

        }

        protected void gEntityTypePropertyList_RowDataBound(object sender, GridViewRowEventArgs e)
        {
            if (e.Row != null && e.Row.RowType == DataControlRowType.DataRow)
            {
                var btnRockEntityInfo = e.Row.FindControl("btnRockEntityInfo") as BootstrapButton;
                var btnAddPropertyPath = e.Row.FindControl("btnAddPropertyPath") as BootstrapButton;
                var btnRemovePropertyPath = e.Row.FindControl("btnRemovePropertyPath") as BootstrapButton;
                var lPropertyTypeIcons = e.Row.FindControl("lPropertyTypeIcons") as System.Web.UI.WebControls.Literal;
                var lPropertyDebufInfo = e.Row.FindControl("lPropertyDebufInfo") as System.Web.UI.WebControls.Literal;


                if (btnRockEntityInfo != null)
                {
                    var entityprop = e.Row.DataItem as PropertyPathModel;
                    btnRockEntityInfo.Visible = entityprop.IsRockClass;

                    var buttonbrowsepath = String.IsNullOrEmpty(BrowsingPath) ?
                        btnRockEntityInfo.CommandName = entityprop.CoderName :
                        BrowsingPath + "." + entityprop.CoderName;

                    btnRockEntityInfo.CommandName = buttonbrowsepath;
                    btnRockEntityInfo.CommandArgument = entityprop.TypeName;

                    if (btnAddPropertyPath != null)
                    {
                        btnAddPropertyPath.Visible = !entityprop.Selected;
                        btnAddPropertyPath.CommandArgument = buttonbrowsepath;
                        btnAddPropertyPath.Attributes.Add("TypeName", entityprop.TypeName);
                        btnAddPropertyPath.Attributes.Add("PropertyName", entityprop.Name);
                        btnAddPropertyPath.Attributes.Add("CoderName", entityprop.CoderName);
                        btnAddPropertyPath.Attributes.Add("PropertyType", entityprop.PropertyType);
                        btnAddPropertyPath.Attributes.Add("IsRockClass", entityprop.IsRockClass.ToString());
                        btnAddPropertyPath.Attributes.Add("IsArray", entityprop.IsArray.ToString());
                        btnAddPropertyPath.Attributes.Add("Enabled", entityprop.Enabled.ToString());
                        btnAddPropertyPath.Attributes.Add("Selected", entityprop.Selected.ToString());


                        if (entityprop.IsRockClass)
                        {
                            btnAddPropertyPath.CommandName = "AddEntityPath";

                        }
                        else
                        {
                            btnAddPropertyPath.CommandName = "AddPropertyPath";
                        }
                    }

                    if (btnRemovePropertyPath != null)
                    {
                        btnRemovePropertyPath.Visible = entityprop.Selected;
                        btnRemovePropertyPath.CommandArgument = buttonbrowsepath;
                        btnRemovePropertyPath.Attributes.Add("TypeName", entityprop.TypeName);
                        btnRemovePropertyPath.Attributes.Add("PropertyName", entityprop.Name);
                        btnRemovePropertyPath.Attributes.Add("CoderName", entityprop.CoderName);
                        btnRemovePropertyPath.Attributes.Add("PropertyType", entityprop.PropertyType);
                        btnRemovePropertyPath.Attributes.Add("IsRockClass", entityprop.IsRockClass.ToString());
                        btnRemovePropertyPath.Attributes.Add("IsArray", entityprop.IsArray.ToString());
                        btnRemovePropertyPath.Attributes.Add("Enabled", entityprop.Enabled.ToString());
                        btnRemovePropertyPath.Attributes.Add("Selected", entityprop.Selected.ToString());

                        if (entityprop.IsRockClass)
                        {
                            btnRemovePropertyPath.CommandName = "RemoveEntityPath";
                        }
                        else
                        {
                            btnRemovePropertyPath.CommandName = "RemovePropertyPath";
                        }
                    }


                    if (lPropertyTypeIcons != null)
                    {
                        if (entityprop.IsRockClass && entityprop.IsArray)
                        {
                            lPropertyTypeIcons.Text = "<i class='fa fa-users text-info' data-toggle='tooltip' title='Rock Entity Array'></i>";
                            lPropertyTypeIcons.Text += "  <i class='fab fa-rockrms text-info'></i>";
                        }
                        else if (entityprop.IsRockClass)
                        {
                            lPropertyTypeIcons.Text = "<i class='fa fa-user text-info' data-toggle='tooltip' title='Rock Entity'></i>";
                            lPropertyTypeIcons.Text += "  <i class='fab fa-rockrms text-info'></i>";
                        }
                        else if (entityprop.IsArray)
                        {
                            lPropertyTypeIcons.Text = "<i class='fa fa-list text-info' data-toggle='tooltip' title='Property Array'></i>";
                        }
                        else
                        {
                            lPropertyTypeIcons.Text = "<i class='far fa-calendar-alt text-info' data-toggle='tooltip' title='Property'></i>";
                        }
                    }

                    if (!EntityExport.EntityTypeIsExportable(entityprop.TypeName))
                    {
                        btnRockEntityInfo.Text = "<i class='fa fa-info text-warning' data-toggle='tooltip' title='Forbidden to Export'></i>";
                    }
                    else
                    {
                        btnRockEntityInfo.Text = "<i class='fa fa-info text-info'></i>";
                    }
                }
            }
        }

        protected void tvSelectedEntities_TreeNodeExpanded(object sender, TreeNodeEventArgs e)
        {
            if (e.Node.Value.StartsWith("Rock.Model."))
            {
                var props = EntityProperties.AsQueryable();

                if (EntityProperties.Where(p => p.EntityTypeName == e.Node.Value).Any())
                {
                    props = EntityProperties.Where(p => p.EntityTypeName == e.Node.Value).AsQueryable();
                }
                else
                {
                    var entitytype = EntityTypeCache.Get(e.Node.Value);
                    var reflectionprops = EntityExport.GetEntityTypeProperties(entitytype, UnselectPropertyList, HidePropertyList);
                    props = ConvertEntityPropertiesToViewModel(reflectionprops.ToList()).AsQueryable();
                }

                var propsnode = new TreeNode();
                propsnode.Text = "Properties";
                propsnode.Value = "Properties";

                foreach (var prop in props)
                {
                    var propnode = new TreeNode();
                    propnode.Text = prop.Name;
                    propnode.Value = prop.TypeName;

                    propsnode.ChildNodes.Add(propnode);
                }

                e.Node.Value = "Properties";
                e.Node.ChildNodes.Clear();
                e.Node.ChildNodes.Add(propsnode);
            }
        }
    }

    #region Entity Coder Helper Class

    /// <summary>
    /// Copy of the EntityCoder class used to replicate the Workflow Type coder but for generic entities.
    /// TODO: Need to move this class to a DLL or extensions.
    /// </summary>
    public class EntityCodingHelper : EntityCoder
    {
        public List<QueuedEntity> Entities { get; private set; }

        /// <summary>
        /// Initialize a new Helper object for facilitating the export/import of entities.
        /// </summary>
        /// <param name="rockContext">The RockContext to work in when exporting or importing.</param>
        public EntityCodingHelper(RockContext rockContext) : base()
        {
            Entities = new List<QueuedEntity>();
        }


        /// <summary>
        /// Generate the list of entities that reference this parent entity. These are entities that
        /// must be created after this entity has been created.
        /// </summary>
        /// <param name="parentEntity">The parent entity to find reverse-references to.</param>
        /// <param name="path">The property path that led us to this final property.</param>
        /// <param name="exporter">The object that handles filtering during an export process.</param>
        /// <returns>A list of KeyValuePairs that identify the property names and entites to be followed.</returns>
        public List<PropertyInfo> FindChildEntities(Type parentEntityType)
        {
            var children = new List<PropertyInfo>();

            var properties = parentEntityType.GetProperties()
                    .Where(p => p.Name != "Id" && p.Name != "Guid")
                    .Where(p => p.Name != "ForeignId" && p.Name != "ForeignGuid" && p.Name != "ForeignKey")
                    .ToList();

            //TODO: Copied from EntityCoder
            // Take a stab at any properties that are an ICollection<IEntity> and treat those
            // as child entities.
            //
            foreach (var property in properties)
            {
                if (property.PropertyType.GetInterface("IEnumerable") != null && property.PropertyType.GetGenericArguments().Length == 1)
                {
                    if (typeof(IEntity).IsAssignableFrom(property.PropertyType.GetGenericArguments()[0]))
                    {
                        children.Add(property);
                    }

                    //TODO: There is a better way to do this
                    else if (property.PropertyType.FullName.StartsWith("Rock.Model."))
                    {
                        var modeltype = EntityTypeCache.Get(property.PropertyType.FullName);
                        if (modeltype != null && modeltype.IsEntity && EntityExport.EntityTypeIsExportable(modeltype.Name))
                        {
                            children.Add(property);
                        }
                    }
                }
            }

            return children;
        }



        /// <summary>
        /// Find entities that this object references directly. These are entities that must be
        /// created before this entity can be re-created.
        /// </summary>
        /// <param name="parentEntity">The parent entity whose references we need to find.</param>
        /// <param name="path">The property path that led us to this final property.</param>
        /// <param name="exporter">The object that handles filtering during an export process.</param>
        /// <returns>A dictionary that identify the property names and entites to be followed.</returns>
        public Dictionary<PropertyInfo, PropertyInfo> FindReferencedEntities(Type parentEntityType)
        {
            var references = new Dictionary<PropertyInfo, PropertyInfo>();
            var properties = parentEntityType.GetProperties()
                    .Where(p => p.Name != "Id" && p.Name != "Guid")
                    .Where(p => p.Name != "ForeignId" && p.Name != "ForeignGuid" && p.Name != "ForeignKey")
                    .ToList();

            //
            // Take a stab at any properties that end in "Id" and likely reference another
            // entity, such as a property called "WorkflowId" probably references the Workflow
            // entity and should be linked by Guid.
            //
            foreach (var property in properties)
            {
                if (property.Name.EndsWith("Id") && (property.PropertyType == typeof(int) || property.PropertyType == typeof(Nullable<int>)))
                {
                    var entityProperty = parentEntityType.GetProperty(property.Name.Substring(0, property.Name.Length - 2));

                    if (entityProperty == null || !typeof(IEntity).IsAssignableFrom(entityProperty.PropertyType))
                    {
                        continue;
                    }
                    references.Add(property, entityProperty);
                }
                else
                {
                    if (property.PropertyType.FullName.Length <= 100)
                    {
                        var rocktype = EntityTypeCache.Get(property.PropertyType.FullName);
                        if (rocktype != null && rocktype.IsEntity && property.PropertyType.FullName.Length <= 100)
                        {
                            //Check for a database table decoration on the class. If it has a table declaration treat it as a child entity                            
                            var databasetable = rocktype.GetEntityType().GetType().CustomAttributes.
                                Where(a => a.GetType().FullName == "System.ComponentModel.DataAnnotations.Schema.TableAttribute").FirstOrDefault();

                            var tablename = databasetable != null ? databasetable.GetPropertyValue("Name").ToString() : null;
                            if (!String.IsNullOrEmpty(tablename))
                            {
                                references.Add(property, null);
                            }
                        }
                    }
                }
            }

            return references;
        }
    }

    #endregion




    #region Entity Export Class

    /// <summary>
    /// Defines the rules for exporting a Rock Entity.
    /// This is an attempt to create a generic set of rules that can be used in most cases.
    /// There could be a need to create custom exporters like the Workflow Type Exporter.
    /// Page comes to mind, it could prove useful to prevent circular loops.
    /// </summary>
    /// <seealso cref="EntityCoding.IExporter" />
    public class EntityExport : IExporter
    {
        #region Properties

        public IEntity RootEntity { get; set; }

        #endregion

        #region Rock Entity Methods

        public static List<EntityProperty> GetEntityTypeProperties(EntityTypeCache entitytype, string[] unselectedlist, string[] hidelist)
        {
            var properties = new List<EntityProperty>();
            PropertyInfo propertyTypeInfo = null;
            PropertyInfo matchedPropertyInfo = null;

            var props = entitytype.GetEntityType().GetProperties()
                     .Where(p => System.Attribute.IsDefined(p, typeof(DataMemberAttribute)))
                     .Where(p => !System.Attribute.IsDefined(p, typeof(NotMappedAttribute)))
                     .Where(p => !System.Attribute.IsDefined(p, typeof(DatabaseGeneratedAttribute)))
                     //.Where(p => p.Name != "Id" && p.Name != "Guid")
                     //.Where(p => p.Name != "ForeignId" && p.Name != "ForeignGuid" && p.Name != "ForeignKey")
                     //.Where(p => p.Name != "CreatedByPersonAliasId" && p.Name != "ModifiedByPersonAliasId")
                     .ToList();



            using (var context = new RockContext())
            {
                var coder = new EntityCodingHelper(context);
                var childprops = coder.FindChildEntities(entitytype.GetEntityType());
                var relatedprops = coder.FindReferencedEntities(entitytype.GetEntityType());


                foreach (var prop in props)
                {
                    var propertyinfo = new EntityProperty();
                    propertyinfo.Namespace = prop.PropertyType.Namespace;
                    propertyinfo.ProcessingMessages = new List<string>();
                    propertyinfo.PropertyTypeName = prop.PropertyType.FullName;
                    propertyinfo.Name = prop.Name;
                    propertyinfo.IsRockClass = childprops.Contains(prop) || relatedprops.ContainsKey(prop);
                    propertyTypeInfo = prop;
                    propertyinfo.LavaHidden = System.Attribute.IsDefined(prop, typeof(LavaHiddenAttribute));
                    propertyinfo.IsNotMapped = System.Attribute.IsDefined(prop, typeof(NotMappedAttribute));
                    propertyinfo.IsRequired = System.Attribute.IsDefined(prop, typeof(RequiredAttribute))
                        || prop.Name == "Id" || prop.Name == "Guid";

                    propertyinfo.IsDataMember = System.Attribute.IsDefined(prop, typeof(DataMemberAttribute));
                    propertyinfo.IsArray = (typeof(IEnumerable).IsAssignableFrom(prop.PropertyType) && prop.PropertyType != typeof(string));

                    if (propertyinfo.IsArray)
                    {
                        propertyinfo.ResolvedTypeName = prop.PropertyType.GenericTypeArguments.Last().FullName;
                        propertyinfo.ResolvedTypeFullName = prop.PropertyType.GenericTypeArguments.Last().FullName;
                    }

                    //Match the database property to the entity
                    if (relatedprops.ContainsKey(prop))
                    {

                        if (prop.Name == "CreatedByPersonAlias" || prop.Name == "CreatedByPersonAliasId")
                        {
                            string s = "";
                        }

                        if (relatedprops.ContainsKey(prop) && relatedprops[prop] != null)
                        {
                            propertyinfo.ProcessingMessages.Add("The named property was found and matched to the Id column");
                            propertyinfo.ExportPropertyType = "Related Entity";
                            propertyinfo.DatabaseColumnName = prop.Name;
                            propertyinfo.Name = relatedprops[prop].Name;
                            propertyinfo.LavaHidden =
                                System.Attribute.IsDefined(relatedprops[prop], typeof(LavaHiddenAttribute)) &&
                                System.Attribute.IsDefined(prop, typeof(LavaHiddenAttribute));
                            propertyinfo.IsNotMapped =
                                System.Attribute.IsDefined(relatedprops[prop], typeof(NotMappedAttribute)) &&
                                System.Attribute.IsDefined(prop, typeof(NotMappedAttribute));
                            propertyinfo.IsRequired =
                                System.Attribute.IsDefined(relatedprops[prop], typeof(RequiredAttribute)) ||
                                System.Attribute.IsDefined(prop, typeof(RequiredAttribute));

                            propertyinfo.IsArray = prop is IEnumerable || relatedprops[prop] is IEnumerable;
                            propertyinfo.ResolvedTypeName = relatedprops[prop].PropertyType.FullName;
                            propertyinfo.ResolvedTypeFullName = relatedprops[prop].PropertyType.FullName;
                            matchedPropertyInfo = relatedprops[prop];


                        }
                        else
                        {
                            propertyinfo.ProcessingMessages.Add("The Id column was not found for the related entity type");
                            propertyinfo.ExportPropertyType = "Related Entity No Id";

                        }
                    }

                    if (childprops.Contains(prop))
                    {
                        propertyinfo.ExportPropertyType = "Child Entity Set";
                        if (prop.PropertyType.GenericTypeArguments.Any())
                        {
                            propertyinfo.ResolvedTypeName = prop.PropertyType.GenericTypeArguments.Last().FullName;
                            propertyinfo.ResolvedTypeFullName = prop.PropertyType.GenericTypeArguments.Last().FullName;
                        }
                    }


                    if (System.Attribute.IsDefined(prop, typeof(DataMemberAttribute)) && String.IsNullOrEmpty(propertyinfo.DatabaseColumnName))
                    {
                        propertyinfo.DatabaseColumnName = prop.Name;
                    }

                    if (String.IsNullOrEmpty(propertyinfo.ResolvedTypeName))
                    {
                        if (prop.PropertyType.Namespace == "Rock.Model")
                        {
                            propertyinfo.ResolvedTypeName = prop.PropertyType.FullName;
                        }
                        else if (prop.PropertyType.GenericTypeArguments.Any())
                        {
                            propertyinfo.ResolvedTypeName = prop.PropertyType.GenericTypeArguments.Last().Name;
                            propertyinfo.ResolvedTypeFullName = prop.PropertyType.GenericTypeArguments.Last().FullName;
                        }
                        else
                        {
                            propertyinfo.ResolvedTypeName = prop.PropertyType.Name;
                            propertyinfo.ResolvedTypeFullName = prop.PropertyType.FullName;
                        }
                    }

                    foreach (var entityprop in properties.Where(p => p.ExportPropertyType == "Related Entity"))
                    {
                        var matchinginfo = properties.Where(
                            p => p.Name == entityprop.Name
                            && p.ExportPropertyType != "Related Entity").FirstOrDefault();
                        if (matchinginfo != null && !String.IsNullOrEmpty(matchinginfo.Name))
                        {
                            entityprop.ResolvedTypeName = matchinginfo.ResolvedTypeName;
                        }
                    }

                    properties.Add(propertyinfo);
                }
            }


            properties = properties.Where(p => !properties.Where(x => matchedPropertyInfo != null && matchedPropertyInfo == propertyTypeInfo).Any()).ToList();

            foreach (var entityprop in properties)
            {
                if (ForbiddenEntityTypes.Where(f => f.Name == entityprop.ResolvedTypeName).Any()
                    && !String.IsNullOrEmpty(entityprop.ResolvedTypeFullName)
                    && entityprop.ResolvedTypeFullName.StartsWith("Rock")
                    )
                {
                    entityprop.Enabled = false;
                }
                else
                {
                    entityprop.Enabled = !unselectedlist.Contains(entityprop.DatabaseColumnName);
                    entityprop.Selected =
                        !unselectedlist.Contains(entityprop.DatabaseColumnName) &&
                        !hidelist.Contains(entityprop.DatabaseColumnName);
                }

                entityprop.Visible = !hidelist.Contains(entityprop.DatabaseColumnName);
            }

            return properties.OrderByDescending(p => p.ExportPropertyType).ThenBy(p => p.Name).ToList();
        }

        /// <summary>
        /// This method contains a list of banned entity types.
        /// We don't want to export any personal information or large amounts of data.
        /// </summary>
        /// <param name="path"></param>
        /// <returns>False if the entity type should never be exported.</returns>
        public static bool EntityTypeIsExportable(string entitytypename)
        {
            if (String.IsNullOrEmpty(entitytypename))
            {
                return false;
            }
            try
            {
                var sysnames = new List<string>() {
                    "System.String",
                    "System.DateTime",
                    "System.Guid",
                    "System.Int32",
                    "System.Boolean" };

                var systype = Type.GetType(entitytypename);
                if (systype != null)
                {
                    if (sysnames.Contains(systype.Name))
                    {
                        return true;
                    }
                    else if (systype.BaseType != null)
                    {
                        if (sysnames.Contains(systype.BaseType.Name))
                        {
                            return true;
                        }
                        if (systype.BaseType.BaseType != null)
                        {
                            if (sysnames.Contains(systype.BaseType.BaseType.Name))
                            {
                                return true;
                            }
                        }
                    }
                }


                if (systype != null && sysnames.Contains(systype.Name))
                {
                    return true;
                }
            }
            catch (Exception ex)
            {
                return false;
            }

            //Rock core entity types should never be imported or exported
            if (entitytypename == "Rock.Model.EntityType" || entitytypename == "Rock.Model.PersonAlias")
            {
                return false;
            }

            try
            {

                if (String.IsNullOrEmpty(entitytypename) || entitytypename.Length > 100)
                {
                    var type = Type.GetType(entitytypename);
                    if (type != null && type.BaseType != null && type.BaseType.Name == "System.ValueType")
                    {
                        return true;
                    }
                    return false;
                }
            }
            catch (Exception ex)
            {
                return false;
            }

            //Explicitly allowed entity types
            //Edge cases that don't follow the rule, explicitly include these configuration entities for security and performance
            if (

                //Other financial tables are all excluded
                entitytypename == "Rock.Model.FinancialStatementTemplate" ||

                //Including some group tables in case we want to export security or something that is stored in a group
                entitytypename == "Rock.Model.Group" ||
                entitytypename == "Rock.Model.GroupLocation" ||
                entitytypename == "Rock.Model.GroupLocationSchedule" ||
                entitytypename == "Rock.Model.GroupLocationScheduleConfig" ||
                entitytypename == "Rock.Model.GroupRequirement" ||
                entitytypename == "Rock.Model.GroupRequirementType" ||
                entitytypename == "Rock.Model.GroupScheduleExclusion" ||
                entitytypename == "Rock.Model.GroupSync" ||
                entitytypename == "Rock.Model.GroupType" ||
                entitytypename == "Rock.Model.GroupTypeAssociation" ||
                entitytypename == "Rock.Model.GroupTypeLocationType" ||
                entitytypename == "Rock.Model.GroupTypeRole" ||

                //Other entity types that would be excluded by the blanket filter
                entitytypename == "Rock.Model.InteractionChannel" ||
                entitytypename == "Rock.Model.InteractionComponent" ||
                entitytypename == "Rock.Model.InteractionDeviceType" ||
                entitytypename == "Rock.Model.NoteType" ||
                entitytypename == "Rock.Model.NotificationMessageType" ||
                entitytypename == "Rock.Model.PersonalizationSegment" ||
                entitytypename == "Rock.Model.PersonalizedEntity" ||
                entitytypename == "Rock.Model.RegistrationInstance" ||
                entitytypename.StartsWith("Rock.Model.RegistrationTemplate", StringComparison.OrdinalIgnoreCase) ||
                entitytypename == "Rock.Model.Reminder" ||
                entitytypename == "Rock.Model.RemoteAuthenticationSession" ||
                entitytypename == "Rock.Model.SignatureDocument" ||
                entitytypename == "Rock.Model.Step" ||
                entitytypename == "Rock.Model.StepWorkflow" ||
                entitytypename == "Rock.Model.Streak" ||
                entitytypename == "Rock.Model.TaggedItem" ||
                entitytypename == "Rock.Model.SignatureDocument" ||
                entitytypename == "Rock.Model.Workflow" ||
                entitytypename == "Rock.Model.WorkflowAction" ||
                entitytypename == "Rock.Model.WorkflowActivity"
               )
            {
                return true;
            }

            //Blocked entity types blanket filter
            //Exclude chunks of tables if they start with certain strings
            if (
                entitytypename.StartsWith("Rock.Model.Analytics", StringComparison.OrdinalIgnoreCase) ||
                entitytypename.StartsWith("Rock.Model.Attendance", StringComparison.OrdinalIgnoreCase) ||
                entitytypename.StartsWith("Rock.Model.Auth", StringComparison.OrdinalIgnoreCase) ||
                entitytypename.StartsWith("Rock.Model.Financial", StringComparison.OrdinalIgnoreCase) ||
                entitytypename.StartsWith("Rock.Model.BackgroundCheck", StringComparison.OrdinalIgnoreCase) ||
                entitytypename.StartsWith("Rock.Model.Following", StringComparison.OrdinalIgnoreCase) ||

                //Don't allow exporting of group member, group location etc. Only the group itself and scheduling info.
                entitytypename.StartsWith("Rock.Model.Group", StringComparison.OrdinalIgnoreCase) ||
                entitytypename.StartsWith("Rock.Model.History", StringComparison.OrdinalIgnoreCase) ||
                entitytypename.StartsWith("Rock.Model.Interaction", StringComparison.OrdinalIgnoreCase) ||

                //Don't allow person data to be exported.
                entitytypename.StartsWith("Rock.Model.Person", StringComparison.OrdinalIgnoreCase) ||
                entitytypename.StartsWith("Rock.Model.Meta", StringComparison.OrdinalIgnoreCase) ||
                entitytypename.StartsWith("Rock.Model.Note", StringComparison.OrdinalIgnoreCase) ||
                entitytypename.StartsWith("Rock.Model.Notification", StringComparison.OrdinalIgnoreCase) ||

                //Don't allow user data to be exorted.
                entitytypename.StartsWith("Rock.Model.User", StringComparison.OrdinalIgnoreCase)
               )
            {
                return false;
            }

            //Blocked entity types
            //Individuale entity types to exclude
            if (
                entitytypename == "Rock.Model.AchievementAttempt" ||
                entitytypename == "Rock.Model.AdaptiveMessage" ||
                entitytypename == "Rock.Model.Assessment" ||
                entitytypename == "Rock.Model.BenevolenceRequest" ||
                entitytypename == "Rock.Model.BinaryFile" ||
                entitytypename == "Rock.Model.BinaryFileData" ||
                entitytypename == "Rock.Model.Communication" ||
                entitytypename == "Rock.Model.CommunicationRecipient" ||
                entitytypename == "Rock.Model.CommunicationResponse" ||
                entitytypename == "Rock.Model.CommunicationResponseAttachment" ||
                entitytypename == "Rock.Model.ConnectionRequest" ||
                entitytypename == "Rock.Model.ConnectionRequestActivity" ||
                entitytypename == "Rock.Model.ConnectionRequestWorkflow" ||
                entitytypename == "Rock.Model.Device" ||
                entitytypename == "Rock.Model.DeviceLocation" ||
                entitytypename == "Rock.Model.Document" ||
                entitytypename == "Rock.Model.FieldType" ||
                entitytypename == "Rock.Model.GroupMember" ||
                entitytypename == "Rock.Model.MetricValue" ||
                entitytypename == "Rock.Model.MetricValuePartition" ||
                entitytypename == "Rock.Model.NcoaHistory" ||
                entitytypename == "Rock.Model.Note" ||
                entitytypename == "Rock.Model.PluginMigration" || //TODO: Allow these for export? Version history?
                entitytypename == "Rock.Model.PrayerRequest" ||
                entitytypename == "Rock.Model.Registration"
               )
            {
                return false;
            }

            var entitytype = EntityTypeCache.Get(entitytypename);

            if (entitytype != null)
            {
                //If we found the entity type make sure it is a Rock entity and IsEntity = true
                return (entitytype.IsEntity && entitytype.Name.StartsWith("Rock.Model") && entitytype.IsEntity);
            }

            //If it is a plugin entity type or IsEntity = false then don't allow export
            return false;
        }

        /// <summary>
        /// Determines if the entity at the given path requires a new Guid value when it's imported
        /// onto the target system. On import, if an entity of that type and Guid already exists then
        /// it is not imported and a reference to the existing entity is used instead.
        /// </summary>
        /// <param name="path">The path to the queued entity object that is being checked.</param>
        /// <returns>
        ///   <c>true</c> if the path requires a new Guid value; otherwise, <c>false</c>
        /// </returns>
        public bool DoesPathNeedNewGuid(EntityPath path)
        {
            return false;
        }

        /// <summary>
        /// Gets any custom references for the entity at the given path.
        /// </summary>
        /// <param name="parentEntity">The entity that will later be encoded.</param>
        /// <param name="path">The path to the parent entity.</param>
        /// <returns>
        /// A collection of references that should be applied to the encoded entity.
        /// </returns>
        public ICollection<Reference> GetUserReferencesForPath(IEntity parentEntity, EntityPath path)
        {
            if (path == "")
            {
                if (parentEntity.TypeName == "Rock.Model.Page")
                {
                    return new List<Reference>() {
                            Reference.UserDefinedReference( "ParentPageId", "ParentPage" )
                        };
                }
                else if (parentEntity.TypeName == "Rock.Model.Site")
                {
                    return new List<Reference>() {
                            Reference.UserDefinedReference( "LayoutId", "PageLayout" )
                        };
                }

            }

            return null;
        }

        /// <summary>
        /// Determines whether the path to an entity should be considered critical. A critical
        /// entity is one that MUST exist on the target system in order for the export/import to
        /// succeed, as such a critical entity is always included.
        /// </summary>
        /// <param name="path">The path to the queued entity object that is being checked.</param>
        /// <returns>
        ///   <c>true</c> if the path is critical; otherwise, <c>false</c>.
        /// </returns>
        public bool IsPathCritical(EntityPath path)
        {
            return (DoesPathNeedNewGuid(path));
        }

        /// <summary>
        /// Determines if the current entity being processed was the root entity or not.            
        /// </summary>
        /// <param name="path"></param>
        /// <returns></returns>
        public bool IsRootEntity(EntityPath path)
        {
            var entityinfo = path.Last();
            return RootEntity.TypeName == entityinfo.Entity.TypeName && RootEntity.Id == entityinfo.Entity.Id;
        }

        private bool PathIsRequired(EntityPath path)
        {
            return true;
        }


        /// <summary>
        /// Determines if the property at the given path should be followed to it's referenced entity.
        /// This is called for both referenced entities and child entities.
        /// </summary>
        /// <param name="path">The path.</param>
        /// <returns></returns>
        public bool ShouldFollowPathProperty(EntityPath path)
        {
            return false;
        }

        #endregion



        #region Entity Processing Properties

        /// <summary>
        /// Gets a list of entity types not allowed to be exported.
        /// </summary>
        public static List<string> ExportableEntityTypes
        {
            get
            {
                List<string> result = new List<string>();
                foreach (var entitytype in EntityTypeCache.All())
                {
                    if (EntityExport.EntityTypeIsExportable(entitytype.Name))
                    {
                        result.Add(entitytype.Name);
                    }
                }
                return result;
            }
        }

        /// <summary>
        /// Gets a list of entity types not allowed to be exported.
        /// </summary>
        public static List<EntityTypeCache> ForbiddenEntityTypes
        {
            get
            {
                var result = EntityTypeCache.All().Where(t => !EntityExport.EntityTypeIsExportable(t.Name)).ToList();
                if (result.Any())
                {
                    string s = "";
                }
                return result;
            }
        }

        #endregion

    }

    #region Export Helper Classes

    public class EntityProperty
    {
        public bool CaptureForeignKeys { get; set; }
        public bool CaptureEntityDateTimes { get; set; }

        public string Name { get; set; }
        public string ExportPropertyType { get; set; }
        public string DatabaseColumnName { get; set; }
        public bool IsRockClass { get; set; }
        public bool IsRequired { get; set; }
        public bool IsArray { get; set; }
        public bool LavaHidden { get; set; }
        public bool IsDataMember { get; set; }
        public bool IsNotMapped { get; set; }
        public string PropertyTypeName { get; set; }
        public string Namespace { get; set; }
        public string ResolvedTypeName { get; set; }
        public string ResolvedTypeFullName { get; set; }
        public List<string> ProcessingMessages { get; set; }
        public bool Selected { get; set; }
        public bool Enabled { get; set; }
        public bool Visible { get; set; }

    }

    #endregion



    #endregion



}
