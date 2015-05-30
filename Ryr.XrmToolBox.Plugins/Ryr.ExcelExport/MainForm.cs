using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml;
using System.Xml.Linq;
using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Metadata;
using Microsoft.Xrm.Sdk.Query;
using MsCrmTools.ViewLayoutReplicator.Helpers;
using OfficeOpenXml;
using OfficeOpenXml.FormulaParsing.Excel.Functions.RefAndLookup;
using OfficeOpenXml.Style;
using OfficeOpenXml.Table;
using Tanguy.WinForm.Utilities.DelegatesHelpers;
using XrmToolBox;
using XrmToolBox.Attributes;

[assembly: BackgroundColor("")]
[assembly: PrimaryFontColor("")]
[assembly: SecondaryFontColor("Gray")]
[assembly: SmallImageBase64("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")]
[assembly: BigImageBase64("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")]
namespace Ryr.ExcelExport
{
    public partial class MainForm : PluginBase
    {
        private List<EntityMetadata> entitiesCache;
        private Dictionary<string, string>  optionsetCache = new Dictionary<string, string>();
        public MainForm()
        {
            InitializeComponent();
        }

        private void tsbLoadEntities_Click(object sender, EventArgs e)
        {
            ExecuteMethod(LoadEntities);
        }

        private void LoadEntities()
        {
            lvEntities.Items.Clear();
            gbEntities.Enabled = false;
            tsbLoadEntities.Enabled = false;
            tsbRefresh.Enabled = false;
            tsbExportExcel.Enabled = false;
            lvViews.Items.Clear();
            txtFetchXml.Text = "";

            WorkAsync("Loading entities...",
                e =>
                {
                    e.Result = MetadataHelper.RetrieveEntities(Service);
                },
                e =>
                {
                    if (e.Error != null)
                    {
                        string errorMessage = CrmExceptionHelper.GetErrorMessage(e.Error, true);
                        CommonDelegates.DisplayMessageBox(ParentForm, errorMessage, "Error", MessageBoxButtons.OK,
                                                          MessageBoxIcon.Error);
                    }
                    else
                    {
                        entitiesCache = (List<EntityMetadata>)e.Result;
                        lvEntities.Items.Clear();
                        var list = new List<ListViewItem>();
                        foreach (EntityMetadata emd in (List<EntityMetadata>)e.Result)
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
                });
        }

        private void lvEntities_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (lvEntities.SelectedItems.Count > 0)
            {
                string entityLogicalName = lvEntities.SelectedItems[0].Tag.ToString();

                // Reinit other controls
                lvViews.Items.Clear();
                txtFetchXml.Text = "";

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
            tsbExportExcel.Enabled = true;
        }

        private void tsbExportExcel_Click(object sender, EventArgs e)
        {
            var dialog = new SaveFileDialog
            {
                Filter = "Excel XML Spreadsheet (*.xml)|*.xml",
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
            WorkAsync("Retrieving records..", 
                (w, e) =>
                {
                    var outputFile = new ExcelPackage();
                    var ws = outputFile.Workbook.Worksheets.Add("Result");

                    if (lvViews.SelectedItems.Count == 0)
                        return;

                    ListViewItem item = lvViews.SelectedItems[0];
                    var view = (Entity)item.Tag;

                    var fetchXml = view["fetchxml"].ToString();
                    var attributes = XElement.Parse(fetchXml)
                        .Descendants("attribute")
                        .Select(x => x.Attribute("name").Value).ToList();

                    var fetchToQuery = new FetchXmlToQueryExpressionRequest { FetchXml = string.Format(txtFetchXml.Text, 1) };
                    var retrieveQuery = ((FetchXmlToQueryExpressionResponse)Service.Execute(fetchToQuery)).Query;
                    retrieveQuery.PageInfo = new PagingInfo { PageNumber = 1 };

                    var rowNumber = 1;
                    var columnNumber = 1;
                    EntityCollection results;
                    var recordCount = 0;
                    var pageNumber = 0;

                    foreach (var attribute in attributes)
                    {
                        var attributeResponse =
                            (RetrieveAttributeResponse)Service.Execute(new RetrieveAttributeRequest
                            {
                                LogicalName = attribute,
                                EntityLogicalName = retrieveQuery.EntityName
                            });
                        ws.Cells[rowNumber, columnNumber].Value = attributeResponse.AttributeMetadata.DisplayName.UserLocalizedLabel.Label;
                        columnNumber++;
                    }
                    rowNumber++;
                    do
                    {
                        results = Service.RetrieveMultiple(retrieveQuery);
                        w.ReportProgress(0,string.Format("Processing Page {0}...", ++pageNumber));

                        columnNumber = 1;
                        foreach (var result in results.Entities)
                        {
                            foreach (var attribute in attributes)
                            {
                                ws.Cells[rowNumber, columnNumber].Value = result.Contains(attribute) ?
                                    UnwrapAttribute(attribute, retrieveQuery.EntityName, result[attribute]) : string.Empty;
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

                },
                e => MessageBox.Show(string.Format("{0} records exported", e.Result)),
                e => SetWorkingMessage(e.UserState.ToString()));
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
                var cacheKey = string.Format("{0}:{1}", attributeName, entityName);
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
                        attributeUnwrappedValue = RetrieveOptionsetText(((OptionSetValue)attributeValue).Value, attributeName, entityName);
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
                attributeUnwrappedValue = ((DateTime)attributeValue).ToString("s");
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
    }
}
