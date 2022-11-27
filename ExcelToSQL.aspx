<%@ Page Title="" Language="C#" MasterPageFile="~/Master.Master" AutoEventWireup="true" CodeBehind="ExcelToSQL.aspx.cs" Inherits="Excel_to_SQL_Data_Import.ExcelToSQL" %>

<asp:Content ID="Content1" ContentPlaceHolderID="head" runat="server">

     <%-------  BootBox CDN ---------%>
    <script src="https://cdnjs.cloudflare.com/ajax/libs/bootbox.js/6.0.0/bootbox.min.js" integrity="sha512-oVbWSv2O4y1UzvExJMHaHcaib4wsBMS5tEP3/YkMP6GmkwRJAa79Jwsv+Y/w7w2Vb/98/Xhvck10LyJweB8Jsw==" crossorigin="anonymous" referrerpolicy="no-referrer"></script>

</asp:Content>
<asp:Content ID="Content2" ContentPlaceHolderID="ContentPlaceHolder1" runat="server">
    <div class="container">
       
        <div class="row justify-content-md-center">

            <div class="row col-md-6 ">

                <div class="border border-warning ">
                   
                    <div class="bd-callout bd-callout-warning">
                        <h5 id="conveying-meaning-to-assistive-technologies">Excel Sheet To SQL Insert Data ASP.NET C#</h5>
                         <div class="col-md-6 ">

                               <div class="row col-md-12 text-right">
                                   <asp:LinkButton Text="Download Excel Demo" id="lbDownload" OnClick="lbDownload_Click"  runat="server" />

                                  </div>

                             <label>Upload Excel Sheet</label>

                             <asp:FileUpload CssClass="form-control" ID="FileUploadExcel" runat="server" />

                             <asp:Button ID="btnUploadExcel" OnClick="btnUploadExcel_Click" CssClass="form-control " BackColor="#ffff66" BorderStyle="None"   Text="Upload Excel" runat="server" />
                             </div>
                       
                    </div>
                </div>
            </div>
            </div>
        </div>
          
    
        
</asp:Content>
