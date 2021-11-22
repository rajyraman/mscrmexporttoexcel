using McTools.Xrm.Connection;
using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Metadata;
using Microsoft.Xrm.Sdk.Query;
using MsCrmTools.ViewLayoutReplicator.Helpers;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Windows.Forms;
using System.Xml.Linq;
using Tanguy.WinForm.Utilities.DelegatesHelpers;
using XrmToolBox.Extensibility;
using XrmToolBox.Extensibility.Args;
using XrmToolBox.Extensibility.Interfaces;

namespace Ryr.ExcelExport
{
    public partial class ExcelExportPlugin : PluginControlBase, IMessageBusHost, IGitHubPlugin, IHelpPlugin, IStatusBarMessenger, IShortcutReceiver
    {
        private static int fileNumber = 0;
        private List<EntityMetadata> entitiesCache;
        private string fetchXml;
        private Dictionary<string, object> optionsetCache = new Dictionary<string, object>();

        public ExcelExportPlugin()
        {
            InitializeComponent();
        }

        private void tsbLoadEntities_Click(object sender, EventArgs e)
        {
            ExecuteMethod(LoadEntities);
        }

        public override void UpdateConnection(IOrganizationService newService, ConnectionDetail detail, string actionName, object parameter)
        {
            base.UpdateConnection(newService, detail, actionName, parameter);
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
                        lvEntities.AutoResizeColumns(ColumnHeaderAutoResizeStyle.ColumnContent);
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
            lvViews.AutoResizeColumns(ColumnHeaderAutoResizeStyle.ColumnContent);
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
                FileName = $"{lvEntities.SelectedItems[0].SubItems[0].Text}-{DateTime.Today.ToString("yyyyMMdd")}.xlsx"
            };
            if (dialog.ShowDialog() == DialogResult.OK)
            {
                ExecuteMethod(ExportCurrentViewToExcel, dialog.FileName);
            }
        }

        private void WriteToExcel(List<string> headers, List<List<string>> rows, string fileName)
        {
            ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
            using (var outputFile = new ExcelPackage())
            {
                var ws = outputFile.Workbook.Worksheets.Add("Result");

                //write headers first
                for (var columnNumber = 0; columnNumber < headers.Count; columnNumber++)
                {
                    ws.Cells[1, columnNumber + 1].Value = headers[columnNumber];
                }

                //actual data rows
                for (var rowNumber = 0; rowNumber < rows.Count; rowNumber++)
                {
                    for (var columnNumber = 0; columnNumber < headers.Count; columnNumber++)
                    {
                        ws.Cells[rowNumber + 2, columnNumber + 1].Value = rows[rowNumber][columnNumber];
                    }
                }

                if (!rows.Any()) return;

                outputFile.File = fileNumber > 0 ?
                    new FileInfo(fileName.Replace(".xlsx", $"_{fileNumber}.xlsx")) :
                    new FileInfo(fileName);
                fileNumber++;
                outputFile.Save();
            }
        }

        private void ExportCurrentViewToExcel(string fileName)
        {
#if DEBUG
            Debugger.Launch();
#endif
            WorkAsync(new WorkAsyncInfo("Retrieving records..", (w, e) =>
            {
                if (lvViews.SelectedItems.Count == 0 || txtFetchXml.Text == string.Empty)
                {
                    return;
                }
                var fetchElements = XElement.Parse(txtFetchXml.Text);
                foreach (var linkElement in fetchElements.Descendants("link-entity"))
                {
                    if (linkElement.Attribute("alias") == null)
                    {
                        linkElement.SetAttributeValue("alias", "e" + string.Join("", Guid.NewGuid().ToString().Split('-')));
                    }
                }

                var fetchToQuery = new FetchXmlToQueryExpressionRequest { FetchXml = fetchElements.ToString() };
                var retrieveQuery = ((FetchXmlToQueryExpressionResponse)Service.Execute(fetchToQuery)).Query;
                var fetchXmlCount = (int)batchSize.Value;
                if(fetchElements.Attribute("count") != null)
                    fetchXmlCount = Convert.ToInt32(fetchElements.Attribute("count")?.Value ?? "500");
                var fetchXmlPageNumber = Convert.ToInt32(fetchElements.Attribute("page")?.Value ?? "1");
                retrieveQuery.PageInfo = new PagingInfo
                {
                    PageNumber = fetchXmlPageNumber,
                    Count = fetchXmlCount
                };
                retrieveQuery.NoLock = true;
                long totalRecordCountForEntity = 0;
                if (ConnectionDetail.OrganizationMajorVersion >= 9 && ConnectionDetail.OrganizationMinorVersion >= 1)
                {
                    totalRecordCountForEntity = ((RetrieveTotalRecordCountResponse)Service.Execute(
                        new RetrieveTotalRecordCountRequest
                        {
                            EntityNames = new string[] { retrieveQuery.EntityName }
                        })).EntityRecordCountCollection.First().Value;
                }
                fileNumber = 0;
                EntityCollection results;
                var processedRecordCount = 0;
                var pageNumber = 0;
                var headers = new List<string>();
                var rows = new List<List<string>>();

                var attributes = RetrieveAttributeMetadata(fetchElements, w, headers);

                do
                {
                    results = Service.RetrieveMultiple(retrieveQuery);
                    if (totalRecordCountForEntity > 0)
                    {
                        w.ReportProgress(0, $"Processing Page {++pageNumber}, {results.Entities.Count}/{totalRecordCountForEntity} ({Math.Round(results.Entities.Count * 100.0 / totalRecordCountForEntity, 2)}%) records...");
                    }
                    else
                    {
                        w.ReportProgress(0, $"Processing Page {++pageNumber}, {results.Entities.Count} records. Total processed {processedRecordCount}...");
                    }

                    foreach (var result in results.Entities)
                    {
                        var rowValues = new List<string>();
                        foreach (var attribute in attributes)
                        {
                            var attributeName = string.IsNullOrEmpty(attribute.Alias)
                                ? attribute.AttributeName
                                : $"{attribute.Alias}.{attribute.AttributeName}";

                            var attributeValue = string.Empty;
                            if (result.Contains(attributeName))
                            {
                                var unWrappedAttributeValue = UnwrapAttribute(attribute.AttributeName, attribute.EntityName,
                                    result[attributeName]);
                                if (unWrappedAttributeValue != null)
                                {
                                    attributeValue = unWrappedAttributeValue.ToString();
                                }
                            }
                            rowValues.Add(attributeValue);
                        }

                        rows.Add(rowValues);

                        if (rows.Count == maxRowsPerFile.Value)
                        {
                            WriteToExcel(headers, rows, fileName);
                            rowValues.Clear();
                            rows.Clear();
                        }
                    }
                    processedRecordCount += results.Entities.Count;
                    retrieveQuery.PageInfo.PageNumber++;
                    retrieveQuery.PageInfo.PagingCookie = results.PagingCookie;

                    if (fetchElements.Attribute("count") != null) break; //if count is specified in FetchXml don't page
                } while (results.MoreRecords);

                //Write any leftover rows
                if (rows.Any())
                {
                    WriteToExcel(headers, rows, fileName);
                }
                e.Result = processedRecordCount;

            })
            {
                PostWorkCallBack = (c) =>
                {
                    if (c.Error != null)
                    {
                        MessageBox.Show(this, "An error occured while exporting the Excel file: " + c.Error.ToString(), "Error",
                                        MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                    else
                    {
                        SendMessageToStatusBar(this, new StatusBarMessageEventArgs($"{c.Result} records exported.."));
                        if (DialogResult.Yes == MessageBox.Show(this, "Do you want to open exported file?", "Question", MessageBoxButtons.YesNo, MessageBoxIcon.Question))
                        {
                            Process.Start(fileName);
                        }
                    }
                },
                ProgressChanged = (c) => SetWorkingMessage(c.UserState.ToString())
            });
        }

        private List<AttributeInfo> RetrieveAttributeMetadata(XElement fetchElements, BackgroundWorker w, List<string> headers)
        {
            var attributes = fetchElements
                .Descendants("attribute")
                .Select(x => new AttributeInfo
                {
                    AttributeName = x.Attribute("name").Value,
                    EntityName = x.Parent.Attribute("name").Value,
                    Alias = x.Parent.Attribute("alias") != null ? x.Parent.Attribute("alias").Value : string.Empty,
                }).ToList();

            foreach (var attribute in attributes)
            {
                w.ReportProgress(0, $"Retrieving metadata for {attribute.AttributeName}...");
                var attributeResponse =
                    (RetrieveAttributeResponse)Service.Execute(new RetrieveAttributeRequest
                    {
                        LogicalName = attribute.AttributeName,
                        EntityLogicalName = attribute.EntityName
                    });
                if (attributeResponse.AttributeMetadata.DisplayName.UserLocalizedLabel == null &&
                    attributeResponse.AttributeMetadata.AttributeOf != null)
                {
                    w.ReportProgress(0,
                        $"Retrieving metadata for {attributeResponse.AttributeMetadata.AttributeOf}...");
                    attributeResponse =
                        (RetrieveAttributeResponse)Service.Execute(new RetrieveAttributeRequest
                        {
                            LogicalName = attributeResponse.AttributeMetadata.AttributeOf,
                            EntityLogicalName = attribute.EntityName
                        });
                }
                var headerValue = attributeResponse.AttributeMetadata.DisplayName.UserLocalizedLabel != null ?
                    attributeResponse.AttributeMetadata.DisplayName.UserLocalizedLabel.Label : attributeResponse.AttributeMetadata.LogicalName;
                headers.Add(headerValue);
            }
            return attributes;
        }

        private object UnwrapAttribute(string attributeName, string entityName, object attributeValue)
        {
            if (attributeValue == null)
            {
                return string.Empty;
            }
            string cacheKey;
            switch (attributeValue)
            {
                case EntityReference e:
                    return e.Name;
                case OptionSetValue o:
                    var optionSetValue = o.Value;
                    return RetrieveOptionsetText(optionSetValue, attributeName, entityName);
                case OptionSetValueCollection oc:
                    var optionSetValueCollection = oc;
                    var optionSetValueTextValues = new List<string>();
                    foreach(var optionSet in oc)
                    {
                        optionSetValueTextValues.Add(RetrieveOptionsetText(optionSet.Value, attributeName, entityName));
                    }                                       
                    return string.Join("; ", optionSetValueTextValues);
                case bool b:
                    cacheKey = $"{attributeName}:{entityName}:{b}";
                    if (!optionsetCache.ContainsKey(cacheKey))
                    {
                        optionsetCache[cacheKey] = RetrieveBooleanLabel((bool)attributeValue, attributeName, entityName);
                    }
                    return optionsetCache[cacheKey];
                case Money m:
                    return m.Value;
                case Guid g:
                    return g.ToString("B");
                case AliasedValue a:
                    return UnwrapAttribute(attributeName, entityName, a.Value);
                case DateTime d:
                    return d.Kind == DateTimeKind.Utc ? d.ToLocalTime().ToString(CultureInfo.CurrentCulture.DateTimeFormat) : d.ToString(CultureInfo.CurrentCulture.DateTimeFormat);
                default:
                    return attributeValue;
            }
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
            if (optionSets is BooleanAttributeMetadata b)
            {
                var optionMetaData = optionsetValue ? b.OptionSet.TrueOption : b.OptionSet.FalseOption;
                optionsetText = optionMetaData.Label.UserLocalizedLabel.Label;
            }
            return optionsetText;
        }        

        private string RetrieveOptionsetText(int optionsetValue, string attributeName, string entityName)
        {
            var cacheKey = $"{attributeName}:{entityName}:{optionsetValue}";
            if (!optionsetCache.ContainsKey(cacheKey))
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
                if (optionSets is EnumAttributeMetadata p)
                {
                    optionMetaData = p.OptionSet.Options.FirstOrDefault(x => x.Value == optionsetValue);
                }
                if (optionMetaData != null)
                {
                    optionsetText = optionMetaData.Label.UserLocalizedLabel.Label;
                }
                return optionsetText;
            }
            else
            {
                return optionsetCache[cacheKey].ToString();
            }
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
            messageBusEventArgs.TargetArgument = fetchXml;
            OnOutgoingMessage(this, messageBusEventArgs);
        }

        public void OnIncomingMessage(MessageBusEventArgs message)
        {
            if (message.SourcePlugin != "FetchXML Builder") return;

            txtFetchXml.Text = (string)message.TargetArgument;
            FormatXML(true);
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

        public void ReceiveKeyDownShortcut(KeyEventArgs e)
        {
            if (e.KeyCode == Keys.F5 && tsbExportExcel.Enabled)
            {
                RequestFileDetails();
            }
        }

        public void ReceiveKeyPressShortcut(KeyPressEventArgs e)
        {
        }

        public void ReceiveKeyUpShortcut(KeyEventArgs e)
        {
        }

        public void ReceivePreviewKeyDownShortcut(PreviewKeyDownEventArgs e)
        {
        }

        public event EventHandler<MessageBusEventArgs> OnOutgoingMessage;
        public event EventHandler<StatusBarMessageEventArgs> SendMessageToStatusBar;

        public string RepositoryName => "mscrmexporttoexcel";

        public string UserName => "rajyraman";

        public string HelpUrl => "https://github.com/rajyraman/mscrmexporttoexcel/blob/master/README.md";
    }

    internal class AttributeInfo
    {
        public string AttributeName { get; set; }
        public string EntityName { get; set; }
        public string Alias { get; set; }
    }
}
