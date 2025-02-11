﻿<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="home.aspx.cs" Inherits="test.home" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
<meta http-equiv="Content-Type" content="text/html; charset=utf-8"/>
    <title></title>
    <!-- 引入 Bootstrap 5 -->
    <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.2/dist/css/bootstrap.min.css" rel="stylesheet"/>
    <link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/bootstrap-icons/font/bootstrap-icons.css"/>
    <script src="https://cdn.jsdelivr.net/npm/bootstrap@5.3.2/dist/js/bootstrap.bundle.min.js"></script>

</head>
<body>
    <form id="form1" runat="server">
        <nav class="navbar navbar-expand-lg navbar-light bg-light">
            <div class="container-fluid"><!--<div class="container-fluid d-flex justify-content-end">-->
                <div class="ms-auto">
                    <asp:Button runat="server" Text="Logout" OnClick="Logout_Click" CssClass="btn btn-danger" style="float:right;"/>
                </div>
            </div>
        </nav>
        <div class="container mt-4">
            <div>
                <asp:Panel ID="Nav_Panel" runat="server" style="padding:7px;margin-bottom:5px;">
                    <asp:Button ID="AddPanel_btn" runat="server" Text="Add" OnClick="AddPanel_btn_Click" CssClass="btn btn-primary" style="margin-bottom:5px;"/>
                    <asp:TextBox ID="Search_text" runat="server"></asp:TextBox>
                    <asp:Button ID="Search_btn" runat="server" Text="Search" OnClick="Search_Click" CssClass="btn btn-primary" style="margin-bottom:5px;"/>
                    <asp:Label ID="message1" runat="server" ForeColor="#FF5050" Visible="False"></asp:Label>
                    <asp:Button ID="ShowAll_btn" runat="server" Text="ShowAll" OnClick="ShowAll_Click" CssClass="btn btn-primary" Style="margin-bottom: 5px;" Visible="False" />
                </asp:Panel>

                <asp:Panel ID="Add_Panel" runat="server" Visible="False" Class="border  border-3 rounded" style="padding:7px;margin-bottom:5px;">
                    <div class="card-body text-center">
                        <div class="row g-3">
                            <div class="col-12">
                                <div class="input-group">
                                    <span class="input-group-text"><i class="bi bi-info-circle"></i></span>
                                    <asp:TextBox ID="AddName_text" runat="server" class="form-control" style="background-color:white" placeholder="username"></asp:TextBox>
                                </div>
                            </div>
                            <div class="col-12">
                                <div class="input-group">
                                    <span class="input-group-text"><i class="bi bi-person"></i></span>
                                    <asp:TextBox ID="AddAccount_text" runat="server" class="form-control" style="background-color:white" placeholder="account"></asp:TextBox>
                                </div>
                            </div>
                            <div class="col-12">
                                <div class="input-group">
                                    <span class="input-group-text"><i class="bi bi-lock"></i></span>
                                    <asp:TextBox ID="AddPassword_text" runat="server" class="form-control" style="background-color:white" placeholder="password"></asp:TextBox><br />
                                </div>
                            </div>
                            <div class="col-12">
                                <div class="input-group">
                                    <span class="input-group-text"><i class="bi bi-telephone"></i></span>
                                    <asp:TextBox ID="AddPhone_text" runat="server" class="form-control" style="background-color:white" placeholder="phome"></asp:TextBox><br />
                                </div>
                            </div>
                            <div class="col-12">
                                <div class="input-group">
                                    <span class="input-group-text"><i class="bi bi-balloon"></i></span>
                                    <asp:DropDownList ID="AddGender_list" runat="server" class="form-control" style="background-color:white">
                                        <asp:ListItem>male</asp:ListItem>
                                        <asp:ListItem>female</asp:ListItem>
                                    </asp:DropDownList>
                                </div>
                            </div>
                        </div>
                    </div>
                    <br />
                    <asp:Button ID="add_btn" runat="server" Text="add" OnClick="Add_Click" class="btn btn-success" ValidationGroup="add" />&nbsp
                    <asp:Button ID="cancel_btn" runat="server" Text="cancel" OnClick="Cancel_Click" class="btn btn-secondary"/>
                    <asp:Label ID="message" runat="server" ForeColor="#FF5050"></asp:Label>
                    <asp:RequiredFieldValidator ID="RequiredFieldValidator1" runat="server" ErrorMessage="請輸入姓名" ForeColor="#FF5050" ControlToValidate="AddName_text" ValidationGroup="add"></asp:RequiredFieldValidator>
                    <asp:RequiredFieldValidator ID="RequiredFieldValidator2" runat="server" ErrorMessage="請輸入帳號" ForeColor="#FF5050" ControlToValidate="AddAccount_text" ValidationGroup="add"></asp:RequiredFieldValidator>
                    <asp:RequiredFieldValidator ID="RequiredFieldValidator3" runat="server" ErrorMessage="請輸入密碼" ForeColor="#FF5050" ControlToValidate="AddPassword_text" ValidationGroup="add"></asp:RequiredFieldValidator>
                    <br />
                    <asp:RegularExpressionValidator ID="RegularExpressionValidator1" runat="server" ErrorMessage="電話格式錯誤" ControlToValidate="AddPhone_text" ValidationExpression="((\d{10})|(((\(\d{2}\))|(\d{2}-))?\d{4}(-)?\d{3}(\d)?))" ForeColor="#FF5050" ValidationGroup="add"></asp:RegularExpressionValidator>
                </asp:Panel>
            </div>

            <asp:GridView ID="GridView1" runat="server" CssClass="table table-bordered table-primary text-white table-hover" OnRowCommand="GridView1_RowCommand">
                <Columns>
                     <asp:ButtonField ButtonType="Button" Text="delete" CommandName="deletey"/>
                     <asp:ButtonField ButtonType="Button" Text="edit" CommandName="edity" />
                </Columns>
                <Columns>
                     <asp:TemplateField ShowHeader="False">
                                <EditItemTemplate>
                                    <asp:Button ID="btnOk" runat="server" OnClick="btnOk_Click" Text="確定" Style="min-width: 38px; height: 25px; font-size: 15px" CssClass="select-btn" />
                                    <asp:Button ID="btnCel" runat="server" OnClick="btnCel_Click" Text="取消" BackColor="Gray" Style="min-width: 38px; height: 25px; font-size: 15px" CssClass="select-btn" />
                                </EditItemTemplate>
                            </asp:TemplateField>
                </Columns>
                <HeaderStyle CssClass="table-warning text-white" />
                <RowStyle CssClass="table-light" />
            </asp:GridView>   
            <div class="d-flex justify-content-between mt-3">
                <asp:Button ID="Export" runat="server" Text="Export" OnClick="Export_Click" CssClass="btn btn-primary"/>
                <asp:Button ID="Import" runat="server" Text="Import" OnClick="Import_Click" CssClass="btn btn-primary"/>
            </div>
        </div>
    </form>
</body>
</html>
