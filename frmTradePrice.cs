using BMS.Business;
using BMS.Model;
using DevExpress.XtraEditors;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using Outlook = Microsoft.Office.Interop.Outlook;
using System.IO;
using System.Diagnostics;
using DevExpress.Utils;
using BMS.Utils;

namespace BMS
{
    public partial class frmTradePrice : _Forms
    {
        //int employeeID;
        DataSet ds = new DataSet();
        public frmTradePrice()
        {
            InitializeComponent();
        }
        private void frmTradePrice_Load(object sender, EventArgs e)
        {
            LoadProject();
            LoadCustomer();
            LoadEmployee();
            loadData();
        }
        void loadData()
        {
            int empId = -1;
            if (Global.DepartmentID != 1) empId = Global.EmployeeID;

            int employeeId = TextUtils.ToInt(cboEmployee.EditValue);
            int saleAdminId = TextUtils.ToInt(cboEmployeeSaleAdmin.EditValue);
            int projectId = TextUtils.ToInt(cboProject.EditValue);
            int customerId = TextUtils.ToInt(cboCustomer.EditValue);
            string keyword = txtFilterText.Text.Trim();

            ds = TextUtils.LoadDataSetFromSP("spGetTradePrice",
                                                new string[] { "@ID", "@EmpID", "@EmployeeID", "@SaleAdminID", "@ProjectID", "@CustomerID", "@Keyword" },
                                                new object[] { 0, empId, employeeId, saleAdminId, projectId, customerId, keyword });
            grdData.DataSource = ds.Tables[1];
        }
        void loadDetail()
        {
            int empId = -1;
            if (Global.DepartmentID != 1) empId = Global.EmployeeID;

            int id = TextUtils.ToInt(grvData.GetFocusedRowCellValue(colID));
            ds = TextUtils.LoadDataSetFromSP("spGetTradePrice", new string[] { "@ID", "@EmpID" }, new object[] { id, empId });
            TreeData.DataSource = ds.Tables[2];
            TreeData.ExpandAll();
        }

        void LoadProject()
        {
            List<ProjectModel> list = SQLHelper<ProjectModel>.FindAll().OrderByDescending(x => x.ID).ToList();
            cboProject.Properties.DisplayMember = "ProjectName";
            cboProject.Properties.ValueMember = "ID";
            cboProject.Properties.DataSource = list;
        }

        void LoadCustomer()
        {
            //DataTable dt = TextUtils.Select("SELECT ID,CustomerName FROM dbo.Customer where IsDeleted <> 1 Order By CreatedDate DESC");

            var exp1 = new Expression("IsDeleted", 1, "<>");
            var listCustomers = SQLHelper<CustomerModel>.FindByExpression(exp1).OrderByDescending(x => x.CreatedDate).ToList();
            cboCustomer.Properties.DisplayMember = "CustomerName";
            cboCustomer.Properties.ValueMember = "ID";
            cboCustomer.Properties.DataSource = listCustomers;
        }

        void LoadEmployee()
        {
            DataTable dt = TextUtils.LoadDataFromSP("spGetEmployee", "A", new string[] { "@Status" }, new object[] { -1 });
            cboEmployee.Properties.ValueMember = "ID";
            cboEmployee.Properties.DisplayMember = "FullName";
            cboEmployee.Properties.DataSource = dt;

            cboEmployeeSaleAdmin.Properties.ValueMember = "ID";
            cboEmployeeSaleAdmin.Properties.DisplayMember = "FullName";
            cboEmployeeSaleAdmin.Properties.DataSource = dt;

            cboEmployeeSaleAdmin.EditValue = Global.EmployeeID;
        }

        private void btnNew_Click(object sender, EventArgs e)
        {
            frmCalculateTradePrices frm = new frmCalculateTradePrices();
            if (frm.ShowDialog() == DialogResult.OK)
            {
                loadData();
                loadDetail();
                grvData_FocusedRowChanged(null, null);
            }
        }

        private void grvData_FocusedRowChanged(object sender, DevExpress.XtraGrid.Views.Base.FocusedRowChangedEventArgs e)
        {
            loadDetail();
            txtEXW.Text = "EXW = " + TextUtils.ToDecimal(grvData.GetRowCellValue(grvData.FocusedRowHandle, colEXW)).ToString("#,##0.##");
            txtMargin.Text = "Margin = " + TextUtils.ToDecimal(grvData.GetRowCellValue(grvData.FocusedRowHandle, colMargin)).ToString("#,##0.##");
            txtTotalCMPerSet.Text = "Tổng CM/Set = " + TextUtils.ToDecimal(grvData.GetRowCellValue(grvData.FocusedRowHandle, colTotalCMPerSET)).ToString("#,##0.##");
            txtTotalProfit.Text = "Lợi nhuận = " + TextUtils.ToDecimal(grvData.GetRowCellValue(grvData.FocusedRowHandle, colTotalProfit)).ToString("#,##0.##");
            txtTotalProfitRate.Text = "Tỷ lệ lợi nhuận = " + TextUtils.ToDecimal(grvData.GetRowCellValue(grvData.FocusedRowHandle, colTotalProfitPercentFooter)).ToString("#,##0.##") + "%";
        }

        private void btnEdit_Click(object sender, EventArgs e)
        {
            var focusedRowHandle = grvData.FocusedRowHandle;
            if (grvData.RowCount <= 0) return;
            int id = TextUtils.ToInt(grvData.GetFocusedRowCellValue(colID));
            if (id == 0) return;
            TradePriceModel model = (TradePriceModel)TradePriceBO.Instance.FindByPK(id);
            frmCalculateTradePrices frm = new frmCalculateTradePrices();
            frm.tradePrice = model;
            if (frm.ShowDialog() == DialogResult.OK)
            {
                loadData();
                grvData.FocusedRowHandle = focusedRowHandle;
                grvData_FocusedRowChanged(null, null);
            }
        }

        private void btnDelete_Click(object sender, EventArgs e)
        {
            if (grvData.RowCount > 0)
            {
                var focusedRowHandle = grvData.FocusedRowHandle;
                int ID = TextUtils.ToInt(grvData.GetFocusedRowCellValue(colID));
                if (ID == 0) return;
                List<TradePriceDetailModel> list = SQLHelper<TradePriceDetailModel>.FindByAttribute("TradePriceID", ID);
                if (MessageBox.Show(string.Format("Bạn có muốn xóa phiếu hay không ?"), TextUtils.Caption, MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                {
                    TradePriceBO.Instance.Delete(ID);
                    if (list.Count > 0)
                    {
                        TradePriceDetailBO.Instance.DeleteByAttribute("TradePriceID", ID);
                        TreeData.DeleteSelectedNodes();
                    }
                    TradePriceBO.Instance.Delete(ID);
                    grvData.DeleteSelectedRows();
                    grvData.FocusedRowHandle = focusedRowHandle;
                    grvData_FocusedRowChanged(null, null);
                }
            }
        }

        private void updateSaleStatus(int status)
        {
            string isApproveText = "";
            if (status == 1)
            {
                isApproveText = "Chốt";
            }
            if (status == 2)
            {
                isApproveText = "Hủy chốt";
            }
            if (status == 3)
            {
                isApproveText = "Yêu cầu duyệt";
            }
            DialogResult dialog = MessageBox.Show($"Bạn có chắc muốn {isApproveText} danh sách dự án đã chọn không?", "Thông báo", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            if (dialog != DialogResult.Yes) return;
            int[] listSelectedRow = grvData.GetSelectedRows();
            if (listSelectedRow.Length <= 0)
            {
                MessageBox.Show("Vui lòng chọn dự án muốn cập nhật!", "Thông báo");
                return;
            }
            int receiverMailID = 0;
            bool isSendEmail = false;
            //string[] arrBody = new string[listSelectedRow.Length];
            List<string> listBody = new List<string>();
            if (status == 3)
            {
                frmSelectMailReceiver frm = new frmSelectMailReceiver();
                frm.ShowDialog();
                receiverMailID = frm.receiverID;
                if (frm.DialogResult == DialogResult.OK && receiverMailID != 0)
                {
                    isSendEmail = true;
                }
                else return;
            }
            foreach (var row in listSelectedRow)
            {
                //  string projectName = TextUtils.ToString(grvData.GetRowCellValue(row, colProjectName));
                //  string customerName = TextUtils.ToString(grvData.GetRowCellValue(row, colCustomerName));
                //  DateTime saleRequestDate = TextUtils.ToDate(grvData.GetRowCellValue(row, colSaleApprovedDate).ToString());
                int id = TextUtils.ToInt(grvData.GetRowCellValue(row, colID));
                int leaderStatusID = TextUtils.ToInt(grvData.GetRowCellValue(row, colLeaderStatusID));
                int saleStatusID = TextUtils.ToInt(grvData.GetRowCellValue(row, colSaleStatusID));
                string projectCode = TextUtils.ToString(grvData.GetRowCellValue(row, colProjectCode));
                string customerCode = TextUtils.ToString(grvData.GetRowCellValue(row, colCustomerCode));
                int saleAdminID = TextUtils.ToInt(grvData.GetRowCellValue(row, colSaleAdminID));
                int saleEmpID = TextUtils.ToInt(grvData.GetRowCellValue(row, colEmployeeID));
                decimal totalProfitPercent = TextUtils.ToDecimal(grvData.GetRowCellValue(row, colTotalProfitPercent));
                string formatTotalProfitPercent = totalProfitPercent.ToString("0.000");
                decimal totalCM = TextUtils.ToDecimal(grvData.GetRowCellValue(row, colTotalCM));
                string formatTotalCM = totalCM.ToString("0.000");

                //trạng thái hiện tại là yc duyệt nhưng muốn sửa trạng thái chốt
                //if (saleStatusID == 3 && status != 3)
                //{
                //    MessageBox.Show($"Dự án {projectCode} đã yêu cầu duyệt. Không thể chỉnh sửa trạng thái chốt!", TextUtils.Caption, MessageBoxButtons.OK, MessageBoxIcon.Information);
                //    continue;
                //}

                if (saleStatusID == 3 && status != 3 && leaderStatusID != 2)
                {
                    MessageBox.Show($"Dự án {projectCode} đang chờ duyệt. Không thể chỉnh sửa trạng thái chốt!", TextUtils.Caption, MessageBoxButtons.OK, MessageBoxIcon.Information);
                    continue;
                }

                // trạng thái chưa chốt nhưng muốn cập nhật thành yc duyệt
                if (status == 3 && (saleStatusID == 2 || saleStatusID == 0))
                {
                    MessageBox.Show($"Dự án {projectCode} chưa được chốt. Không thể yêu cầu duyệt!", TextUtils.Caption, MessageBoxButtons.OK, MessageBoxIcon.Information);
                    continue;
                }
                // trạng thái đã yc duyệt và tiếp tục cập nhật thành yc duyệt
                if (status == 3 && saleStatusID == 3)
                {
                    MessageBox.Show($"Dự án {projectCode} đã được yêu cầu duyệt trước đó!", TextUtils.Caption, MessageBoxButtons.OK, MessageBoxIcon.Information);
                    continue;
                }
                if (id <= 0) continue;

                TradePriceModel tradePrice = SQLHelper<TradePriceModel>.FindByID(id);
                if (tradePrice == null) return;
                tradePrice.IsApprovedSale = status;
                tradePrice.SaleApprovedDate = DateTime.Now;
                tradePrice.ApprovedSaleID = Global.EmployeeID;
                tradePrice.ApprovedLeaderID = receiverMailID;
                tradePrice.IsApprovedLeader = 0;
                tradePrice.IsApprovedBGD = 0;
                TradePriceBO.Instance.Update(tradePrice);
                if (isSendEmail == true)
                {
                    EmployeeModel saleAdmin = SQLHelper<EmployeeModel>.FindByID(saleAdminID);
                    EmployeeModel saleEmp = SQLHelper<EmployeeModel>.FindByID(saleEmpID);
                    //  string subject = $"Nhân viên {Global.AppFullName} xin duyệt giá cho các khách hàng khách hàng";
                    //   string To = TextUtils.ToString(receiverMailID).Trim();
                    //string cc = "";
                    string body = customerCode + ";" + projectCode + ";" + formatTotalProfitPercent + ";" + formatTotalCM + ";" + saleAdmin.FullName + ";" + saleEmp.FullName;
                    listBody.Add(body);
                }
            }

            if (listBody.Count <= 0) return;
            // Send Email
            EmployeeModel receiver = SQLHelper<EmployeeModel>.FindByID(receiverMailID);
            string subject = $"Nhân viên {Global.AppFullName} xin duyệt giá cho các khách hàng".ToUpper();
            string tbody = "";
            foreach (var item in listBody)
            {
                string[] arrBody = item.Split(';');
                tbody += $"<tr>" +
                    $"<td style=\"border: 1px solid;\">{arrBody[0]}</td>" +
                    $"<td style=\"border: 1px solid;\">{arrBody[1]}</td>" +
                    $"<td style=\"border: 1px solid; text-align: right;\">{TextUtils.ToDecimal(arrBody[2]).ToString("n2")}</td>" +
                    $"<td style=\"border: 1px solid; text-align: right;\">{TextUtils.ToDecimal(arrBody[3]).ToString("n0")}</td>" +
                    $"<td style=\"border: 1px solid;\">{arrBody[4]}</td>" +
                    $"<td style=\"border: 1px solid;\">{arrBody[5]}</td>" +
                    $"</tr>";
            }
            string Body = $"<div>Dear anh/chị {receiver.FullName},<br>" +
                  $"Nhân viên {Global.AppFullName} xin duyệt giá cho các khách hàng <br><br>" +
                  $"<table style=\"border-collapse: collapse;border: 1px solid; table-layout: auto; width: 100%;\"> " +
                  $"<thead>" +
                  $"<tr>" +
                  $"<th style=\"border: 1px solid;\"> Khách hàng </th >" +
                  $"<th style=\"border: 1px solid;\"> Dự án </th>" +
                  $"<th style=\"border: 1px solid;\"> Tổng lợi nhuận dự án (%) </th>" +
                  $"<th style=\"border: 1px solid;\"> Tổng CM (VNĐ )</th>" +
                  $"<th style=\"border: 1px solid;\"> Admin tính giá </th>" +
                  $"<th style=\"border: 1px solid;\"> Sale phụ trách </th>" +
                  $"</tr>" +
                  $"</thead> " +
                  $"<tbody>" + tbody +
                  $" </tbody>" +
                  $"</table> " +
                  $"<br>" +
                  $"Trân trọng,<br>" +
                  $"{Global.AppFullName}." +
                  $"</div>";
            EmailSender.SendEmail(subject, Body, receiverMailID, "");
        }
        private bool updateLeaderStatus(int status)
        {
            int countUpdate = 0;
            string isApproveText = "";
            if (status == 1)
            {
                isApproveText = "Duyệt";
            }
            if (status == 2)
            {
                isApproveText = "Hủy duyệt";
            }
            if (status == 3)
            {
                isApproveText = "Yêu cầu BGĐ duyệt";
            }
            DialogResult dialog = MessageBox.Show($"Bạn có chắc muốn {isApproveText} danh sách dự án đã chọn không?", "Thông báo", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            if (dialog != DialogResult.Yes) return false;


            int[] listSelectedRow = grvData.GetSelectedRows();
            if (listSelectedRow.Length <= 0)
            {
                MessageBox.Show("Vui lòng chọn dự án muốn cập nhật!", "Thông báo");
                return false;
            }

            int receiverMailID = 0;
            bool isSendEmail = false;
            //string[] arrBody = new string[listSelectedRow.Length];
            List<string> listBody = new List<string>();
            if (status == 3)
            {
                frmSelectMailReceiver frm = new frmSelectMailReceiver();
                frm.ShowDialog();
                receiverMailID = frm.receiverID;
                if (frm.DialogResult == DialogResult.OK && receiverMailID != 0)
                {
                    isSendEmail = true;
                }
                else return false;
            }
            foreach (var row in listSelectedRow)
            {
                int id = TextUtils.ToInt(grvData.GetRowCellValue(row, colID));
                //int id = TextUtils.ToInt(grvData.GetRowCellValue(row, colID));
                int saleStatusID = TextUtils.ToInt(grvData.GetRowCellValue(row, colSaleStatusID));
                int leaderStatusID = TextUtils.ToInt(grvData.GetRowCellValue(row, colLeaderStatusID));
                int bgdStatusID = TextUtils.ToInt(grvData.GetRowCellValue(row, colBGDStatusID));

                string projectCode = TextUtils.ToString(grvData.GetRowCellValue(row, colProjectCode));
                string customerCode = TextUtils.ToString(grvData.GetRowCellValue(row, colCustomerCode));
                int saleAdminID = TextUtils.ToInt(grvData.GetRowCellValue(row, colSaleAdminID));
                int saleEmpID = TextUtils.ToInt(grvData.GetRowCellValue(row, colEmployeeID));
                decimal totalProfitPercent = TextUtils.ToDecimal(grvData.GetRowCellValue(row, colTotalProfitPercent));
                string formatTotalProfitPercent = totalProfitPercent.ToString("0.000");
                decimal totalCM = TextUtils.ToDecimal(grvData.GetRowCellValue(row, colTotalCM));
                string formatTotalCM = totalCM.ToString("0.000");

                if (saleStatusID != 3)
                {
                    MessageBox.Show($"Dự án {projectCode} chưa được Sale yêu cầu duyệt!", TextUtils.Caption, MessageBoxButtons.OK, MessageBoxIcon.Information);
                    continue;
                }

                //if (leaderStatusID == 3 && status != 3)
                //{
                //    MessageBox.Show($"Dự án {projectCode} đã được yêu cầu BGĐ duyệt. Không thể chỉnh sửa trạng thái duyệt!", TextUtils.Caption, MessageBoxButtons.OK, MessageBoxIcon.Information);
                //    continue;
                //}
                if (leaderStatusID == 3 && status != 3 && bgdStatusID != 2)
                {
                    MessageBox.Show($"Dự án {projectCode} đang chờ BGĐ duyệt. Không thể chỉnh sửa trạng thái duyệt!", TextUtils.Caption, MessageBoxButtons.OK, MessageBoxIcon.Information);
                    continue;
                }
                // dự án đã được yc duyệt trc đó
                if (leaderStatusID == 3 && status == 3)
                {
                    MessageBox.Show($"Dự án {projectCode} đã được yêu cầu BGĐ duyệt trước đó!", TextUtils.Caption, MessageBoxButtons.OK, MessageBoxIcon.Information);
                    continue;
                }
                if (id <= 0) continue;
                //if (status == 3 && leaderStatusID != 3)
                //{
                //    frmSelectMailReceiver frm = new frmSelectMailReceiver();
                //    frm.ShowDialog();
                //    receiverMailID = frm.receiverID;
                //    if (frm.DialogResult == DialogResult.OK && receiverMailID != 0)
                //    {
                //        sendEmail(projectCode, customerCode, receiverMailID, saleAdminID, saleEmpID, formatTotalProfitPercent, formatTotalCM);

                //    }
                //    else return false;
                //}
                TradePriceModel tradePrice = SQLHelper<TradePriceModel>.FindByID(id);
                if (tradePrice == null) return false;
                tradePrice.IsApprovedLeader = status;
                tradePrice.LeaderApprovedDate = DateTime.Now;
                tradePrice.ApprovedLeaderID = Global.EmployeeID;
                tradePrice.IsApprovedBGD = 0;
                TradePriceBO.Instance.Update(tradePrice);
                countUpdate++;
                if (isSendEmail == true)
                {
                    EmployeeModel saleAdmin = SQLHelper<EmployeeModel>.FindByID(saleAdminID);
                    EmployeeModel saleEmp = SQLHelper<EmployeeModel>.FindByID(saleEmpID);
                    //  string subject = $"Nhân viên {Global.AppFullName} xin duyệt giá cho các khách hàng khách hàng";
                    //   string To = TextUtils.ToString(receiverMailID).Trim();
                    //string cc = "";
                    string body = customerCode + ";" + projectCode + ";" + formatTotalProfitPercent + ";" + formatTotalCM + ";" + saleAdmin.FullName + ";" + saleEmp.FullName;
                    listBody.Add(body);
                }
            }
            if (countUpdate == 0) return false;

            if (listBody.Count > 0 && status == 3)
            {
                // Send Email
                EmployeeModel receiver = SQLHelper<EmployeeModel>.FindByID(receiverMailID);
                string subject = $"Nhân viên {Global.AppFullName} xin duyệt giá cho các khách hàng";
                string tbody = "";
                foreach (var item in listBody)
                {
                    string[] arrBody = item.Split(';');
                    tbody += $"<tr>" +
                        $"<td style=\"border: 1px solid;\">{arrBody[0]}</td>" +
                        $"<td style=\"border: 1px solid;\">{arrBody[1]}</td>" +
                        $"<td style=\"border: 1px solid; text-align: right;\">{TextUtils.ToDecimal(arrBody[2]).ToString("n2")}</td>" +
                        $"<td style=\"border: 1px solid; text-align: right;\">{TextUtils.ToDecimal(arrBody[3]).ToString("n0")}</td>" +
                        $"<td style=\"border: 1px solid;\">{arrBody[4]}</td>" +
                        $"<td style=\"border: 1px solid;\">{arrBody[5]}</td>" +
                        $"</tr>";
                }
                string Body = $"<div>Dear anh/chị {receiver.FullName},<br>" +
                      $"Nhân viên {Global.AppFullName} xin duyệt giá cho các khách hàng <br><br>" +
                      $"<table style=\"border-collapse: collapse;border: 1px solid; table-layout: auto; width: 100%;\"> " +
                      $"<thead>" +
                      $"<tr>" +
                      $"<th style=\"border: 1px solid;\"> Khách hàng </th >" +
                      $"<th style=\"border: 1px solid;\"> Dự án </th>" +
                      $"<th style=\"border: 1px solid;\"> Tổng lợi nhuận dự án (%) </th>" +
                      $"<th style=\"border: 1px solid;\"> Tổng CM (VNĐ )</th>" +
                      $"<th style=\"border: 1px solid;\"> Admin tính giá </th>" +
                      $"<th style=\"border: 1px solid;\"> Sale phụ trách </th>" +
                      $"</tr>" +
                      $"</thead> " +
                      $"<tbody>" + tbody +
                      $" </tbody>" +
                      $"</table> " +
                      $"<br>" +
                      $"Trân trọng,<br>" +
                      $"{Global.AppFullName}." +
                      $"</div>";
                EmailSender.SendEmail(subject, Body, receiverMailID, "");
            }

            return true;
        }
        private void updateBGDStatus(int status, bool isLeaderApprove)
        {
            int[] listSelectedRow = grvData.GetSelectedRows();
            if (isLeaderApprove == false)
            {
                if (listSelectedRow.Length <= 0)
                {
                    MessageBox.Show("Vui lòng chọn dự án muốn cập nhật!", "Thông báo");
                    return;
                }
                string isApproveText = "";
                if (status == 1)
                {
                    isApproveText = "Duyệt";
                }
                if (status == 2)
                {
                    isApproveText = " Hủy duyệt";
                }
                DialogResult dialog = MessageBox.Show($"Bạn có chắc muốn {isApproveText} danh sách dự án đã chọn không?", "Thông báo", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                if (dialog != DialogResult.Yes) return;
            }


            foreach (var row in listSelectedRow)
            {
                int id = TextUtils.ToInt(grvData.GetRowCellValue(row, colID));
                int saleStatusID = TextUtils.ToInt(grvData.GetRowCellValue(row, colSaleStatusID));
                int leaderStatusID = TextUtils.ToInt(grvData.GetRowCellValue(row, colLeaderStatusID));
                string projectCode = TextUtils.ToString(grvData.GetRowCellValue(row, colProjectCode));
                if (saleStatusID != 3)
                {
                    MessageBox.Show($"Dự án {projectCode} chưa được Sale yêu cầu duyệt!", TextUtils.Caption, MessageBoxButtons.OK, MessageBoxIcon.Information);
                    continue;
                }
                if (id <= 0) continue;
                TradePriceModel tradePrice = SQLHelper<TradePriceModel>.FindByID(id);
                if (tradePrice == null) return;
                tradePrice.IsApprovedBGD = status;
                tradePrice.BGDApprovedDate = DateTime.Now;
                tradePrice.ApprovedBGDID = Global.EmployeeID;
                if (leaderStatusID != 3)
                {
                    tradePrice.IsApprovedLeader = status;
                }
                TradePriceBO.Instance.Update(tradePrice);
            }

        }
        private void btnSaleDone_Click(object sender, EventArgs e)
        {
            //int saleStatusID = TextUtils.ToInt(grvData.GetFocusedRowCellValue(colSaleStatusID));
            //string projectCode = TextUtils.ToString(grvData.GetFocusedRowCellValue(colProjectCode));
            //if (saleStatusID == 3)
            //{
            //    updateSaleStatus(1);
            //    loadData();
            //}
            //if (saleStatusID == 2)
            //{
            //    MessageBox.Show($"Dự án {projectCode} đã yêu cầu duyệt. Không thể chỉnh sửa trạng thái chốt!", TextUtils.Caption, MessageBoxButtons.OK, MessageBoxIcon.Information);
            //}
            updateSaleStatus(1);
            loadData();
        }

        private void btnSaleCancelDone_Click(object sender, EventArgs e)
        {
            //int saleStatusID = TextUtils.ToInt(grvData.GetFocusedRowCellValue(colSaleStatusID));
            //string projectCode = TextUtils.ToString(grvData.GetFocusedRowCellValue(colProjectCode));

            //if (saleStatusID == 2)
            //{
            //    MessageBox.Show($"Dự án {projectCode} đã yêu cầu duyệt. Không thể chỉnh sửa trạng thái chốt!", TextUtils.Caption, MessageBoxButtons.OK, MessageBoxIcon.Information);
            //}
            //if (saleStatusID == 1)
            //{
            //    updateSaleStatus(3);
            //    loadData();
            //}
            updateSaleStatus(2);
            loadData();


        }

        private void btnSaleRequest_Click(object sender, EventArgs e)
        {
            //int leaderStatusID = TextUtils.ToInt(grvData.GetFocusedRowCellValue(colLeaderStatusID));
            //string projectCode = TextUtils.ToString(grvData.GetFocusedRowCellValue(colProjectCode));

            //if (leaderStatusID == 1)
            //{

            //    updateSaleStatus(2);
            //    loadData();
            //}
            //else if (leaderStatusID == 2)
            //{
            //    MessageBox.Show($"Dự án {projectCode} đã được yêu cầu duyệt", TextUtils.Caption, MessageBoxButtons.OK, MessageBoxIcon.Information);
            //}
            //else
            //{
            //    MessageBox.Show("Chỉ có thể yêu cầu duyệt các dự án đã chốt!", TextUtils.Caption, MessageBoxButtons.OK, MessageBoxIcon.Information);
            //}
            updateSaleStatus(3);
            loadData();
        }

        private void btnLeaderApprove_Click(object sender, EventArgs e)
        {

            if (updateLeaderStatus(1) == true)
            {
                updateBGDStatus(1, true);
                loadData();
            }

        }

        private void btnLeaderCancel_Click(object sender, EventArgs e)
        {

            //int saleStatusID = TextUtils.ToInt(grvData.GetFocusedRowCellValue(colSaleStatusID));
            //int leaderStatusID = TextUtils.ToInt(grvData.GetFocusedRowCellValue(colLeaderStatusID));
            //if (leaderStatusID == 2)
            //{
            //    MessageBox.Show("Dự án đã được yêu cầu BGĐ duyệt. Không thể chỉnh sửa trạng thái duyệt!", TextUtils.Caption, MessageBoxButtons.OK, MessageBoxIcon.Information);
            //    return;
            //}
            //if (saleStatusID == 2)
            //{
            //    updateLeaderStatus(3);
            //    updateBGDStatus(2);
            //    loadData();
            //}
            //else
            //{
            //    MessageBox.Show("Chỉ có thể hủy duyệt các dự án được Sale yêu cầu!", TextUtils.Caption, MessageBoxButtons.OK, MessageBoxIcon.Information);
            //}
            if (updateLeaderStatus(2) == true)
            {
                updateBGDStatus(2, true);
                loadData();
            }
        }

        private void btnLeaderRequest_Click(object sender, EventArgs e)
        {
            //int saleStatusID = TextUtils.ToInt(grvData.GetFocusedRowCellValue(colSaleStatusID));
            //if (saleStatusID == 2)
            //{
            //    updateLeaderStatus(2);
            //    loadData();
            //}
            //else
            //{
            //    MessageBox.Show("Chỉ có thể yêu cầu duyệt các dự án được Sale yêu cầu!", TextUtils.Caption, MessageBoxButtons.OK, MessageBoxIcon.Information);
            //}
            updateLeaderStatus(3);
            loadData();
        }

        private void btnBGDApprove_Click(object sender, EventArgs e)
        {
            //int saleStatusID = TextUtils.ToInt(grvData.GetFocusedRowCellValue(colSaleStatusID));
            //if (saleStatusID == 2)
            //{
            //    updateBGDStatus(1);
            //    loadData();
            //}
            //else
            //{
            //    MessageBox.Show("Chỉ có thể duyệt các dự án được Sale yêu cầu!", TextUtils.Caption, MessageBoxButtons.OK, MessageBoxIcon.Information);
            //}
            updateBGDStatus(1, false);
            loadData();
        }

        private void btnBGDCancel_Click(object sender, EventArgs e)
        {
            //int saleStatusID = TextUtils.ToInt(grvData.GetFocusedRowCellValue(colSaleStatusID));
            //if (saleStatusID == 2)
            //{
            //    updateBGDStatus(2);
            //    loadData();
            //}
            //else
            //{
            //    MessageBox.Show("Chỉ có thể hủy duyệt các dự án được Sale yêu cầu!", TextUtils.Caption, MessageBoxButtons.OK, MessageBoxIcon.Information);
            //}
            updateBGDStatus(2, false);
            loadData();
        }



        private void sendEmail(string projectCode, string customerCode, int receiverID, int saleAdminID, int employeeID, string totalProfitPercent, string totalCM)
        {
            Outlook.MailItem oMsg;
            Outlook.Application oApp;
            oApp = new Outlook.Application();
            try
            {
                if (receiverID < 0) return;
                EmployeeModel receiver = SQLHelper<EmployeeModel>.FindByID(receiverID);
                EmployeeModel saleAdmin = SQLHelper<EmployeeModel>.FindByID(saleAdminID);
                EmployeeModel saleEmp = SQLHelper<EmployeeModel>.FindByID(employeeID);


                string receiverEmail = receiver.EmailCongTy;
                if (receiver == null) return;

                oMsg = (Outlook.MailItem)oApp.CreateItem(Outlook.OlItemType.olMailItem);
                oMsg.Subject = $"Nhân viên {Global.AppFullName} xin duyệt giá".ToUpper();
                //   oMsg.To = TextUtils.ToString(receiver.EmailCongTy).Trim();
                oMsg.To = TextUtils.ToString(receiver.EmailCongTy).Trim();

                //string cc = "";
                //List<string> lstCC = new List<string>();
                //if (cc.Length > 0)
                //{
                //    string[] arrCC = cc.Split(';');
                //    for (int j = 0; j < arrCC.Length; j++)
                //    {
                //        string mail = arrCC[j];
                //        if (!mail.Contains("@")) continue;
                //        lstCC.Add(arrCC[j]);
                //    }
                //}
                //oMsg.CC = string.Join(";", lstCC);
                //oMsg.Display();
                oMsg.HTMLBody = $"<div><p>Dear {receiver.FullName},<br>" +
                    $"Nhân viên {Global.AppFullName} xin duyệt giá cho khách hàng {customerCode}, dự án {projectCode} <br>" +
                    $"Tổng lợi nhuận dự án {totalProfitPercent}%, CM {totalCM} VNĐ<br>" +
                    $"Admin tính giá {saleAdmin.FullName}, sales phụ trách {saleEmp.FullName}";
                //   $"{Global.AppFullName}.</ p ></ div> ";
                oMsg.Send();
                //oMsg.Close(Outlook.OlInspectorClose.olSave);
                //
                //ExcuteSQL($"UPDATE EmployeeSendEmail SET StatusSend = 2, DateSend = '{DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")}' WHERE ID = {dt.Rows[i]["ID"]}");
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                //string message = $"{DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss")}:\nMessage: {ex.Message}\n{ex.ToString()}\n-------------------------------\n";
                //File.AppendAllText($"logException-{DateTime.Now.ToString("yyyy-MM-dd")}.txt", message);

                //if (ex.Message.Contains("The RPC server is unavailable"))
                //{
                //    IsRun = false;
                //    break;
                //}
            }
            finally
            {
                oMsg = null;
            }
        }

        private void btnExportExcel_Click(object sender, EventArgs e)
        {
            string path = "";
            FolderBrowserDialog fbd = new FolderBrowserDialog();
            if (fbd.ShowDialog() == DialogResult.OK)
            {
                path = fbd.SelectedPath;
            }
            else
            {
                return;
            }
            string fileSourceName = "TINHGIATHUONGMAI.xlsx";

            int id = TextUtils.ToInt(grvData.GetFocusedRowCellValue(colID));
            TradePriceModel tradePrice = SQLHelper<TradePriceModel>.FindByID(id);
            ProjectModel project = SQLHelper<ProjectModel>.FindByID(tradePrice.ProjectID);
            string sourcePath = Application.StartupPath + "\\" + fileSourceName;
            string currentPath = path + "\\" + DateTime.Now.ToString($"_dd_MM_yyyy_HH_mm_ss") + ".xlsx";
            try
            {
                File.Copy(sourcePath, currentPath, true);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Có lỗi khi tạo báo giá!" + Environment.NewLine + ex.Message,
                    TextUtils.Caption, MessageBoxButtons.OK, MessageBoxIcon.Stop);
                return;
            }


            using (WaitDialogForm fWait = new WaitDialogForm("Vui lòng chờ trong giây lát...", "Đang tạo phiếu..."))
            {
                System.Threading.Thread.CurrentThread.CurrentCulture = new System.Globalization.CultureInfo("en-US");
                Excel.Application app = default(Excel.Application);
                Excel.Workbook workBoook = default(Excel.Workbook);
                Excel.Worksheet workSheet = default(Excel.Worksheet);
                try
                {
                    app = new Excel.Application();
                    app.Workbooks.Open(currentPath);
                    workBoook = app.Workbooks[1];
                    workSheet = (Excel.Worksheet)workBoook.Worksheets[1];
                    workSheet.Cells[1, 1] = TextUtils.ToString(project.ProjectCode);
                    workSheet.Cells[3, 1] = TextUtils.ToString(project.ProjectName);
                    workSheet.Cells[4, 25] = TextUtils.ToDecimal(tradePrice.RateCOM);
                    workSheet.Cells[4, 26] = TextUtils.ToDecimal(tradePrice.COM);

                    DataTable dt = (DataTable)TreeData.DataSource;

                    for (int i = 0; i < ds.Tables[0].Columns.Count; i++)
                    {
                        workSheet.Cells[2, 10 + i] = ds.Tables[0].Rows[0][ds.Tables[0].Columns[i].ColumnName];
                    }
                    for (int i = dt.Rows.Count - 1; i >= 0; i--)
                    {
                        workSheet.Cells[8, 2] = TextUtils.ToString(dt.Rows[i]["STT"]);
                        workSheet.Cells[8, 3] = TextUtils.ToString(dt.Rows[i]["Maker"]);
                        workSheet.Cells[8, 4] = TextUtils.ToString(dt.Rows[i]["ProductName"]);
                        workSheet.Cells[8, 5] = TextUtils.ToString(dt.Rows[i]["ProductCode"]); ;
                        workSheet.Cells[8, 6] = TextUtils.ToString(dt.Rows[i]["ProductCodeCustomer"]);
                        workSheet.Cells[8, 7] = TextUtils.ToInt(dt.Rows[i]["Quantity"]);
                        workSheet.Cells[8, 8] = TextUtils.ToString(dt.Rows[i]["Unit"]);
                        workSheet.Cells[8, 9] = TextUtils.ToDecimal(dt.Rows[i]["UnitImportPriceUSD"]);
                        workSheet.Cells[8, 10] = TextUtils.ToDecimal(dt.Rows[i]["TotalImportPriceUSD"]);
                        workSheet.Cells[8, 11] = TextUtils.ToDecimal(dt.Rows[i]["UnitImportPriceVND"]);
                        workSheet.Cells[8, 12] = TextUtils.ToDecimal(dt.Rows[i]["TotalImportPriceVND"]);
                        workSheet.Cells[8, 13] = TextUtils.ToDecimal(dt.Rows[i]["BankCharge"]);
                        workSheet.Cells[8, 14] = TextUtils.ToDecimal(dt.Rows[i]["ProtectiveTariff"]);
                        workSheet.Cells[8, 15] = TextUtils.ToDecimal(dt.Rows[i]["ProtectiveTariffPerPcs"]);
                        workSheet.Cells[8, 16] = TextUtils.ToDecimal(dt.Rows[i]["TotalProtectiveTariff"]);
                        workSheet.Cells[8, 17] = TextUtils.ToDecimal(dt.Rows[i]["OrtherFees"]);
                        workSheet.Cells[8, 18] = TextUtils.ToDecimal(dt.Rows[i]["CustomFees"]);
                        workSheet.Cells[8, 19] = TextUtils.ToDecimal(dt.Rows[i]["TotalImportPriceIncludeFees"]);
                        workSheet.Cells[8, 20] = TextUtils.ToDecimal(dt.Rows[i]["UnitPriceIncludeFees"]);
                        workSheet.Cells[8, 21] = TextUtils.ToDecimal(dt.Rows[i]["CMPerSet"]);
                        workSheet.Cells[8, 22] = TextUtils.ToDecimal(dt.Rows[i]["UnitPriceExpectCustomer"]);
                        workSheet.Cells[8, 23] = TextUtils.ToDecimal(dt.Rows[i]["TotalPriceExpectCustomer"]);
                        workSheet.Cells[8, 24] = TextUtils.ToDecimal(dt.Rows[i]["Profit"]);
                        workSheet.Cells[8, 25] = TextUtils.ToDecimal(dt.Rows[i]["ProfitPercent"]);
                        workSheet.Cells[8, 26] = TextUtils.ToString(dt.Rows[i]["LeadTime"]);
                        workSheet.Cells[8, 29] = TextUtils.ToDecimal(dt.Rows[i]["TotalPrice"]);
                        workSheet.Cells[8, 30] = TextUtils.ToDecimal(dt.Rows[i]["UnitPricePerCOM"]);
                        workSheet.Cells[8, 31] = TextUtils.ToString(dt.Rows[i]["Note"]);

                        ((Excel.Range)workSheet.Rows[8]).Insert();
                    }
                    ((Excel.Range)workSheet.Rows[7]).Delete();
                    ((Excel.Range)workSheet.Rows[7]).Delete();
                    CustomerModel customer = SQLHelper<CustomerModel>.FindByID(tradePrice.CustomerID);
                    string customerName = TextUtils.ToString(customer.CustomerName);
                    string projectCode = TextUtils.ToString(project.ProjectCode);
                    string projectName = TextUtils.ToString(project.ProjectName);
                    string str = string.Join("_", customerName, projectCode, projectName);
                    workSheet.Cells[7, 1] = str;
                    workSheet.Cells[7, 27] = TextUtils.ToDecimal(dt.Rows[0]["TotalPriceLabor"]);
                    workSheet.Cells[7, 28] = TextUtils.ToDecimal(dt.Rows[0]["TotalPriceRTCVision"]);

                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
                finally
                {
                    if (app != null)
                    {
                        app.ActiveWorkbook.Save();
                        app.Workbooks.Close();
                        app.Quit();
                    }
                }
                Process.Start(currentPath);
            }
        }

        private void grdData_DoubleClick(object sender, EventArgs e)
        {
            btnEdit_Click(null, null);
        }

        private void btnImportExcel_Click(object sender, EventArgs e)
        {
            frmTradePriceImportExcel frm = new frmTradePriceImportExcel();

            if (frm.ShowDialog() == DialogResult.OK)
            {
                loadData();
            }
        }

        private void btnFind_Click(object sender, EventArgs e)
        {
            loadData();
        }

        private void btnExportRequest_Click(object sender, EventArgs e)
        {
            int[] listSelectedRow = grvData.GetSelectedRows();
            if (listSelectedRow.Length <= 0)
            {
                MessageBox.Show("Vui lòng chọn sản phẩm muốn xuất kho!", "Thông báo");
                return;
            }
            DialogResult dialog = MessageBox.Show($"Bạn có chắc muốn tạo phiếu xuất kho danh sách dự án đã chọn không?", "Thông báo", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            if (dialog != DialogResult.Yes) return;

            // thông báo list thông báo
            List<string> lsNotificationStatus = new List<string>();
            List<string> lsExportRequested = new List<string>();
            List<string> lsDataNull = new List<string>();
            List<frmBillExportDetail> formList = new List<frmBillExportDetail>();

            // tạo bảng bằng cách clone từ spGetBillExportDetail
            DataTable dtDetail = TextUtils.LoadDataFromSP("spGetBillExportDetail", "A", new string[] { "@BillID" }, new object[] { -100 });
            DataTable dtDetailForForm = dtDetail.Clone();
            dtDetailForForm.Columns.Add("productGroupID");

            foreach (var row in listSelectedRow)
            {
                int isBGDStatus = TextUtils.ToInt(grvData.GetRowCellValue(row, colBGDStatusID));
                string projectCode = TextUtils.ToString(grvData.GetRowCellValue(row, colProjectCode));
                int id = TextUtils.ToInt(grvData.GetRowCellValue(row, colID));
                int projectID = TextUtils.ToInt(grvData.GetRowCellValue(row, colProjectID));
                string projectName = TextUtils.ToString(grvData.GetRowCellValue(row, colProjectName));
                int customerID = TextUtils.ToInt(grvData.GetRowCellValue(row, colCustomerID));
                int saleAdminID = TextUtils.ToInt(grvData.GetRowCellValue(row, colSaleAdminID));
                int productGroupID = 0;

                // kiểm tra xem dữ án đã được BGD duyệt hay chưa
                if (isBGDStatus != 1)
                {
                    lsNotificationStatus.Add(projectCode);
                    continue;
                }

                // gọi dữ liệu chi tiết của bảng TradePriceDetai và sắp xếp theo productGroupID
                //DataSet dtSet = TextUtils.LoadDataSetFromSP("spGetTradePrice", new string[] { "@ID" }, new object[] { id });
                //if (TextUtils.ToInt(dtSet.Tables[2].Rows.Count) == 0)
                //{
                //    lsDataNull.Add(projectCode);
                //    continue;
                //}
                //DataTable dt = dtSet.Tables[2];
                //dt.DefaultView.Sort = "productGroupID ASC";
                //dt = dt.DefaultView.ToTable();

                DataTable dt = TextUtils.Select("SELECT  td.* ,p.ProductCode, p.ProductName, p.ProductNewCode, p.ItemType, p.ProductGroupID, pg.ProductGroupName FROM dbo.TradePriceDetail td LEFT JOIN dbo.ProductSale p ON p.ID = td.ProductID INNER JOIN dbo.ProductGroup pg ON pg.ID = p.ProductGroupID");
                DataTable dtNew = dt.Clone();
                foreach (DataRow rowdt in dt.Rows)
                {
                    if (TextUtils.ToInt(rowdt["TradePriceID"]) == id)
                    {
                        dtNew.ImportRow(rowdt);
                    }
                }
                dtNew.DefaultView.Sort = "productGroupID ASC";
                dtNew = dtNew.DefaultView.ToTable();

                foreach (DataRow dr in dtNew.Rows)
                {
                    int idTrandePriceDetail = TextUtils.ToInt(dr["ID"]);
                    decimal Quantity = TextUtils.ToDecimal(dr["Quantity"]);
                    productGroupID = TextUtils.ToInt(dr["productGroupID"]);

                    List<BillExportDetailModel> listExport = idTrandePriceDetail > 0 ? SQLHelper<BillExportDetailModel>.FindByAttribute("TradePriceDetailID", idTrandePriceDetail) : new List<BillExportDetailModel>();

                    // tính tổng số lượng của BillExportDetail nếu = 0 thì bỏ qua nếu khác = thì tiếp tục
                    if (listExport.Count > 0)
                    {
                        decimal sumQty = listExport.Sum(itemListExport => TextUtils.ToDecimal(itemListExport.Qty));
                        decimal sumNumber = TextUtils.ToDecimal(Quantity - sumQty);
                        Quantity = sumNumber;

                        if (Quantity == 0)
                        {
                            // add thông báo đã có phiếu xuất kho
                            if (!lsExportRequested.Contains(projectCode))
                            {
                                lsExportRequested.Add(projectCode);
                            }
                            continue;
                        }
                    }

                    //nếu productGroupID có sự thay đổi thì add dữ liệu dable vào form mới và xóa dữ liệu bảng đấy đi
                    if (dtDetailForForm.Rows.Count > 0)
                    {
                        int dtProductGroupID = TextUtils.ToInt(dtDetailForForm.Rows[0]["productGroupID"]);
                        if (dtProductGroupID != productGroupID)
                        {
                            AddFormNew(dtDetailForForm, customerID, dtProductGroupID, saleAdminID, formList);
                        }
                    }

                    // add dữ liệu
                    DataRow item = dtDetailForForm.NewRow();
                    item["STT"] = TextUtils.ToInt(dr["STT"]);
                    item["TradePriceDetailID"] = idTrandePriceDetail;
                    item["ProductID"] = TextUtils.ToInt(dr["ProductID"]);
                    item["Qty"] = Quantity;
                    item["ProjectID"] = projectID;
                    item["Note"] = TextUtils.ToString(dr["Note"]);
                    item["ProductCode"] = TextUtils.ToString(dr["ProductCode"]);
                    item["ProductNewCode"] = TextUtils.ToString(dr["ProductNewCode"]);
                    item["ProductName"] = TextUtils.ToString(dr["ProductName"]);
                    item["Unit"] = TextUtils.ToString(dr["Unit"]);
                    item["ProductGroupName"] = TextUtils.ToString(dr["ProductGroupName"]);
                    item["ItemType"] = TextUtils.ToString(dr["ItemType"]);
                    item["ProjectNameText"] = projectName;
                    item["ProjectCodeText"] = projectCode;
                    item["productGroupID"] = productGroupID;
                    dtDetailForForm.Rows.Add(item);
                }

                // tạo form mới nếu là row cuối cùng
                if (dtDetailForForm.Rows.Count > 0)
                {
                    AddFormNew(dtDetailForForm, customerID, productGroupID, saleAdminID, formList);
                }
                else
                {
                    // add thông báo đã có phiếu xuất kho
                    if (!lsExportRequested.Contains(projectCode))
                    {
                        lsExportRequested.Add(projectCode);
                    }
                }
            }

            dtDetailForForm.Clear();

            // Hiển thị các form đã được tạo
            foreach (frmBillExportDetail frm in formList)
            {
                if (frm.ShowDialog() == DialogResult.OK)
                {
                    loadData();
                }
            }

            // hiển thị thông báo
            ShowNotification(lsNotificationStatus, lsExportRequested, lsDataNull);
        }


        private void AddFormNew(DataTable dtDetailForForm, int customerID, int productGroupID, int saleAdminID, List<frmBillExportDetail> formList)
        {
            frmBillExportDetail frmEnd = new frmBillExportDetail();
            frmEnd.customerID = customerID;
            frmEnd.KhoTypeID = productGroupID;
            frmEnd.saleAdminID = saleAdminID;
            frmEnd.dtDetail = dtDetailForForm.Copy();
            frmEnd.WarehouseCode = "HN";
            formList.Add(frmEnd);
            dtDetailForForm.Clear();
        }

        private void ShowNotification(List<string> n1, List<string> n2, List<string> n3)
        {
            string txtNotification = "";

            if (n1.Count > 0)
            {
                txtNotification += "Các mã dự án sau chưa được BGD duyệt:\n";
                txtNotification += string.Join(", ", n1);
            }

            if (n2.Count > 0)
            {
                txtNotification += (txtNotification == "" ? "" : "\n\n") + "Các mã dự án sau đã được yêu cầu xuất kho:\n";
                txtNotification += string.Join(", ", n2);
            }

            if (n3.Count > 0)
            {
                txtNotification += (txtNotification == "" ? "" : "\n\n") + "Các mã dự án sau không có sản phẩm:\n";
                txtNotification += string.Join(", ", n3);
            }

            if (!string.IsNullOrEmpty(txtNotification))
            {
                MessageBox.Show(txtNotification, "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Question);
            }
        }
    }
}
