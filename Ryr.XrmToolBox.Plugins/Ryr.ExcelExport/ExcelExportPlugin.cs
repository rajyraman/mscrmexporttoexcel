using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using System.Xml.Linq;
using McTools.Xrm.Connection;
using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Metadata;
using Microsoft.Xrm.Sdk.Query;
using MsCrmTools.ViewLayoutReplicator.Helpers;
using OfficeOpenXml;
using Tanguy.WinForm.Utilities.DelegatesHelpers;
using Cinteros.Xrm.FetchXmlBuilder;
using XrmToolBox.Extensibility;
using XrmToolBox.Extensibility.Interfaces;

namespace Ryr.ExcelExport
{
    public partial class ExcelExportPlugin : PluginControlBase, IMessageBusHost
    {
        private List<EntityMetadata> entitiesCache;
        private string fetchXml;

        private Dictionary<string, string>  optionsetCache = new Dictionary<string, string>();
        public ExcelExportPlugin()
        {
            InitializeComponent();
        }

        private void tsbLoadEntities_Click(object sender, EventArgs e)
        {
            ExecuteMethod(LoadEntities);
        }

        private void LoadEntities()
        {
            fetchXml = string.Empty;
            lvEntities.Items.Clear();
            gbEntities.Enabled = false;
            tsbLoadEntities.Enabled = false;
            tsbRefresh.Enabled = false;
            tsbExportExcel.Enabled = false;
            tsbEditInFxb.Enabled = false;
            lvViews.Items.Clear();
            txtFetchXml.Text = "";
            WorkAsync(new WorkAsyncInfo("Loading entities...", e =>
            {
                e.Result = MetadataHelper.RetrieveEntities(Service);
            })
            {
                PostWorkCallBack = completedargs =>
                {
                    if (completedargs.Error != null)
                    {
                        string errorMessage = CrmExceptionHelper.GetErrorMessage(completedargs.Error, true);
                        CommonDelegates.DisplayMessageBox(ParentForm, errorMessage, "Error", MessageBoxButtons.OK,
                                                          MessageBoxIcon.Error);
                    }
                    else
                    {
                        entitiesCache = (List<EntityMetadata>)completedargs.Result;
                        lvEntities.Items.Clear();
                        var list = new List<ListViewItem>();
                        foreach (EntityMetadata emd in (List<EntityMetadata>)completedargs.Result)
                        {
                            var item = new ListViewItem { Text = emd.DisplayName.UserLocalizedLabel.Label, Tag = emd.LogicalName };
                            item.SubItems.Add(emd.LogicalName);
                            list.Add(item);
                        }

                        lvEntities.Items.AddRange(list.ToArray());

                        gbEntities.Enabled = true;
                        gbEntities.Enabled = true;
                        tsbLoadEntities.Enabled = true;
                        tsbRefresh.Enabled = true;
                    }
                }
            });
        }

        private void lvEntities_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (lvEntities.SelectedItems.Count > 0)
            {
                string entityLogicalName = lvEntities.SelectedItems[0].Tag.ToString();

                // Reinit other controls
                lvViews.Items.Clear();
                txtFetchXml.Text = string.Empty;
                fetchXml = string.Empty;
                Cursor = Cursors.WaitCursor;

                // Launch treatment
                var bwFillViews = new BackgroundWorker();
                bwFillViews.DoWork += BwFillViewsDoWork;
                bwFillViews.RunWorkerAsync(entityLogicalName);
                bwFillViews.RunWorkerCompleted += BwFillViewsRunWorkerCompleted;
            }
        }

        private void BwFillViewsDoWork(object sender, DoWorkEventArgs e)
        {
            string entityLogicalName = e.Argument.ToString();

            List<Entity> viewsList = ViewHelper.RetrieveViews(entityLogicalName, entitiesCache, Service);
            viewsList.AddRange(ViewHelper.RetrieveUserViews(entityLogicalName, entitiesCache, Service));

            foreach (Entity view in viewsList)
            {
                bool display = true;

                var item = new ListViewItem(view["name"].ToString());
                item.Tag = view;

                #region Gestion de l'image associée à la vue

                switch ((int)view["querytype"])
                {
                    case ViewHelper.VIEW_BASIC:
                        {
                            if (view.LogicalName == "savedquery")
                            {
                                if ((bool)view["isdefault"])
                                {
                                    item.SubItems.Add("Default public view");
                                    item.ImageIndex = 3;
                                }
                                else
                                {
                                    item.SubItems.Add("Public view");
                                    item.ImageIndex = 0;
                                }
                            }
                            else
                            {
                                item.SubItems.Add("User view");
                                item.ImageIndex = 6;
                            }
                        }
                        break;
                    case ViewHelper.VIEW_ADVANCEDFIND:
                        {
                            item.SubItems.Add("Advanced find view");
                            item.ImageIndex = 1;
                        }
                        break;
                    case ViewHelper.VIEW_ASSOCIATED:
                        {
                            item.SubItems.Add("Associated view");
                            item.ImageIndex = 2;
                        }
                        break;
                    case ViewHelper.VIEW_QUICKFIND:
                        {
                            item.SubItems.Add("QuickFind view");
                            item.ImageIndex = 5;
                        }
                        break;
                    case ViewHelper.VIEW_SEARCH:
                        {
                            item.SubItems.Add("Lookup view");
                            item.ImageIndex = 4;
                        }
                        break;
                    default:
                        {
                            //item.SubItems.Add(view["name"].ToString());
                            display = false;
                        }
                        break;
                }

                #endregion

                if (display)
                {
                    // Add view to each list of views (source and target)
                    ListViewItem clonedItem = (ListViewItem)item.Clone();
                    ListViewDelegates.AddItem(lvViews, item);

                    if (view.Contains("iscustomizable") && ((BooleanManagedProperty)view["iscustomizable"]).Value == false
                        && view.Contains("ismanaged") && (bool)view["ismanaged"])
                    {
                        clonedItem.ForeColor = Color.Gray;
                        clonedItem.ToolTipText = "This managed view has not been defined as customizable";
                    }
                }
            }
        }

        private void BwFillViewsRunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            Cursor = Cursors.Default;
            lvViews.Enabled = true;

            if (e.Error != null)
            {
                MessageBox.Show(this, "An error occured: " + e.Error.Message, "Error", MessageBoxButtons.OK,
                                MessageBoxIcon.Error);
            }

            if (lvViews.Items.Count == 0)
            {
                MessageBox.Show(this, "This entity does not contain any view", "Warning", MessageBoxButtons.OK,
                                MessageBoxIcon.Warning);
            }
        }

        private void lvViews_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (lvViews.SelectedItems.Count == 0)
                return;

            ListViewItem item = lvViews.SelectedItems[0];
            var view = (Entity)item.Tag;

            txtFetchXml.Text = XElement.Parse(view["fetchxml"].ToString()).ToString();
            fetchXml = txtFetchXml.Text;
            FormatXML(true);
            tsbExportExcel.Enabled = true;
            tsbEditInFxb.Enabled = true;
        }

        private void tsbExportExcel_Click(object sender, EventArgs e)
        {
            RequestFileDetails();
        }

        private void RequestFileDetails()
        {
            var dialog = new SaveFileDialog
            {
                Filter = "Excel  Workbook(*.xlsx)|*.xlsx",
                FileName = string.Format("{0}-{1}.xlsx",
                    lvEntities.SelectedItems[0].SubItems[0].Text,
                    DateTime.Today.ToString("yyyyMMdd"))
            };
            if (dialog.ShowDialog() == DialogResult.OK)
            {
                ExecuteMethod(ExportCurrentViewToExcel, dialog.FileName);
            }
        }

        private void ExportCurrentViewToExcel(string fileName)
        {
            WorkAsync(new WorkAsyncInfo("Retrieving records..", (w, e) =>
            {
                var outputFile = new ExcelPackage();
                var ws = outputFile.Workbook.Worksheets.Add("Result");

                if (lvViews.SelectedItems.Count == 0 || fetchXml == string.Empty)
                {
                    return;
                }
                var fetchElements = XElement.Parse(fetchXml);
                foreach (var linkElement in fetchElements.Descendants("link-entity"))
                {
                    if (linkElement.Attribute("alias") == null)
                    {
                        linkElement.SetAttributeValue("alias", "e" + string.Join("", Guid.NewGuid().ToString().Split('-')));
                    }
                }
                
                var attributes = fetchElements
                    .Descendants("attribute")
                    .Select(x => new
                    {
                        AttributeName = x.Attribute("name").Value, 
                        EntityName = x.Parent.Attribute("name").Value,
                        Alias = x.Parent.Attribute("alias") != null ? x.Parent.Attribute("alias").Value : string.Empty,
                    }).ToList();

                var fetchToQuery = new FetchXmlToQueryExpressionRequest { FetchXml = fetchElements.ToString() };
                var retrieveQuery = ((FetchXmlToQueryExpressionResponse) Service.Execute(fetchToQuery)).Query;
                retrieveQuery.PageInfo = new PagingInfo { PageNumber = 1 };
                retrieveQuery.PageInfo.Count = fetchElements.Attribute("count") != null ? 
                                                Convert.ToInt32(fetchElements.Attribute("count").Value) : 500;
                var rowNumber = 1;
                var columnNumber = 1;
                EntityCollection results;
                var recordCount = 0;
                var pageNumber = 0;

                foreach (var attribute in attributes)
                {
                    var attributeResponse =
                        (RetrieveAttributeResponse) Service.Execute(new RetrieveAttributeRequest
                        {
                            LogicalName = attribute.AttributeName,
                            EntityLogicalName = attribute.EntityName
                        });
                    ws.Cells[rowNumber, columnNumber].Value =
                        attributeResponse.AttributeMetadata.DisplayName.UserLocalizedLabel.Label;
                    columnNumber++;
                }
                rowNumber++;
                do
                {
                    results = Service.RetrieveMultiple(retrieveQuery);
                    w.ReportProgress(0, string.Format("Processing Page {0}, {1} records...", ++pageNumber, retrieveQuery.PageInfo.Count));

                    columnNumber = 1;
                    foreach (var result in results.Entities)
                    {
                        foreach (var attribute in attributes)
                        {
                            var attributeName = string.IsNullOrEmpty(attribute.Alias)
                                ? attribute.AttributeName
                                : string.Format("{0}.{1}", attribute.Alias, attribute.AttributeName);
                            ws.Cells[rowNumber, columnNumber].Value = result.Contains(attributeName)
                                ? UnwrapAttribute(attribute.AttributeName, attribute.EntityName, result[attributeName])
                                : string.Empty;
                            columnNumber++;
                        }
                        columnNumber = 1;
                        rowNumber++;
                    }
                    recordCount += results.Entities.Count;
                    retrieveQuery.PageInfo.PageNumber++;
                    retrieveQuery.PageInfo.PagingCookie = results.PagingCookie;
                } while (results.MoreRecords);
                ws.Cells[ws.Dimension.Address].AutoFitColumns();

                outputFile.File = new FileInfo(fileName);
                outputFile.Save();
                e.Result = recordCount;

            })
            {
                PostWorkCallBack = (c) => MessageBox.Show(string.Format("{0} records exported", c.Result)),
                ProgressChanged = (c) => SetWorkingMessage(c.UserState.ToString())
            });
        }

        private object UnwrapAttribute(string attributeName, string entityName, object attributeValue)
        {
            object attributeUnwrappedValue;

            if (attributeValue == null)
            {
                return string.Empty;
            }
            if (attributeValue is EntityReference)
            {
                attributeUnwrappedValue = ((EntityReference)attributeValue).Name;
            }
            else
            if (attributeValue is OptionSetValue || attributeValue is bool)
            {
                var optionSetValue = ((OptionSetValue)attributeValue).Value;
                var cacheKey = string.Format("{0}:{1}:{2}", attributeName, entityName, optionSetValue);
                if (optionsetCache.ContainsKey(cacheKey))
                {
                    attributeUnwrappedValue = optionsetCache[cacheKey];
                }
                else
                {
                    if (attributeValue is bool)
                    {
                        attributeUnwrappedValue = RetrieveBooleanLabel((bool)attributeValue, attributeName, entityName);
                    }
                    else
                    {
                        attributeUnwrappedValue = RetrieveOptionsetText(optionSetValue, attributeName, entityName);
                    }
                    optionsetCache.Add(cacheKey, attributeUnwrappedValue.ToString());
                }
            }
            else
            if (attributeValue is Money)
            {
                attributeUnwrappedValue = ((Money)attributeValue).Value;
            }
            else
            if (attributeValue is Guid)
            {
                attributeUnwrappedValue = ((Guid)attributeValue).ToString("B");
            }
            else
            if (attributeValue is DateTime)
            {
                attributeUnwrappedValue = ((DateTime)attributeValue).ToLocalTime().ToString("s");
            }
            else
            if (attributeValue is AliasedValue)
            {
                attributeUnwrappedValue = UnwrapAttribute(attributeName, entityName, ((AliasedValue)attributeValue).Value);
            }
            else
            {
                attributeUnwrappedValue = attributeValue;
            }
            return attributeUnwrappedValue;
        }

        private string RetrieveBooleanLabel(bool optionsetValue, string attributeName, string entityName)
        {
            var optionsetText = string.Empty;
            var retrieveAttributeRequest = new RetrieveAttributeRequest
            {
                EntityLogicalName = entityName,
                LogicalName = attributeName,
                RetrieveAsIfPublished = true
            };
            var retrieveAttributeResponse = (RetrieveAttributeResponse)Service.Execute(retrieveAttributeRequest);
            var optionSets = retrieveAttributeResponse.AttributeMetadata;
            OptionMetadata optionMetaData = null;
            if (optionSets is BooleanAttributeMetadata)
            {
                optionMetaData = optionsetValue
                    ? ((BooleanAttributeMetadata)optionSets).OptionSet.TrueOption
                    : ((BooleanAttributeMetadata)optionSets).OptionSet.FalseOption;
            }
            if (optionMetaData != null)
            {
                optionsetText = optionMetaData.Label.UserLocalizedLabel.Label;
            }
            return optionsetText;
        }

        private string RetrieveOptionsetText(int optionsetValue, string attributeName, string entityName)
        {
            var optionsetText = string.Empty;
            var retrieveAttributeRequest = new RetrieveAttributeRequest
            {
                EntityLogicalName = entityName,
                LogicalName = attributeName,
                RetrieveAsIfPublished = true
            };
            var retrieveAttributeResponse = (RetrieveAttributeResponse)Service.Execute(retrieveAttributeRequest);
            var optionSets = retrieveAttributeResponse.AttributeMetadata;
            OptionMetadata optionMetaData = null;
            if (optionSets is PicklistAttributeMetadata)
            {
                optionMetaData = ((PicklistAttributeMetadata)optionSets).OptionSet.Options.FirstOrDefault(x => x.Value == optionsetValue);
            }
            else if (optionSets is StatusAttributeMetadata)
            {
                optionMetaData = ((StatusAttributeMetadata)optionSets).OptionSet.Options.FirstOrDefault(x => x.Value == optionsetValue);
            }
            else if (optionSets is StateAttributeMetadata)
            {
                optionMetaData = ((StateAttributeMetadata)optionSets).OptionSet.Options.FirstOrDefault(x => x.Value == optionsetValue);
            }
            if (optionMetaData != null)
            {
                optionsetText = optionMetaData.Label.UserLocalizedLabel.Label;
            }
            return optionsetText;
        }

        private void lvEntities_ColumnClick(object sender, ColumnClickEventArgs e)
        {
            lvEntities.SelectedItems.Clear();
            lvEntities.Sorting = lvEntities.Sorting == SortOrder.Ascending ? SortOrder.Descending : SortOrder.Ascending;
            lvEntities.ListViewItemSorter = new ListViewItemComparer(e.Column, lvEntities.Sorting);
        }

        private void lvViews_ColumnClick(object sender, ColumnClickEventArgs e)
        {
            lvViews.SelectedItems.Clear();
            lvViews.Sorting = lvViews.Sorting == SortOrder.Ascending ? SortOrder.Descending : SortOrder.Ascending;
            lvViews.ListViewItemSorter = new ListViewItemComparer(e.Column, lvViews.Sorting);
        }

        private void tsbClose_Click(object sender, EventArgs e)
        {
            base.CloseTool();
        }

        private void tsbRefresh_Click(object sender, EventArgs e)
        {
            ExecuteMethod(LoadEntities);
        }

        private void tsbEditInFxb_Click(object sender, EventArgs e)
        {
            if (lvViews.SelectedItems.Count == 0 || fetchXml == string.Empty)
            {
                MessageBox.Show("No views selected.", "Error");
                return;
            }

            var messageBusEventArgs = new MessageBusEventArgs("FetchXML Builder");
            var fXBMessageBusArgument = new FXBMessageBusArgument(FXBMessageBusRequest.FetchXML)
            {
                FetchXML = fetchXml
            };
            messageBusEventArgs.TargetArgument = fXBMessageBusArgument;
            OnOutgoingMessage(this, messageBusEventArgs);
        }

        public void OnIncomingMessage(MessageBusEventArgs message)
        {

            if (message.SourcePlugin == "FetchXML Builder" &&
                        message.TargetArgument is FXBMessageBusArgument)
            {
                var fxbArg = (FXBMessageBusArgument)message.TargetArgument;
                txtFetchXml.Text = fxbArg.FetchXML;
                FormatXML(true);
                fetchXml = fxbArg.FetchXML;
            }
        }

        private void FormatXML(bool silent)
        {
            try
            {
                txtFetchXml.Process(true);
            }
            catch (Exception ex)
            {
                if (!silent)
                {
                    MessageBox.Show(ex.Message, "XML Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        public event EventHandler<MessageBusEventArgs> OnOutgoingMessage;
    }
}
