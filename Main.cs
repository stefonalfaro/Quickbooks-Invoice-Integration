using Interop.QBFC13;
using System;
using System.Data.SqlClient;
using System.Net;
using System.Net.Sockets;
using System.Threading;
using System.Xml;

namespace QB_queue
{
  internal class Program
  {
    private static SqlConnection connection = new SqlConnection("Data Source=;Initial Catalog=TLSNET;User ID=sa;Password=;Pooling=false;");
    private static SqlConnection connection2 = new SqlConnection("Data Source=;Initial Catalog=TLSNET;User ID=sa;Password=;Pooling=false;");

    private static void Main(string[] args)
    {
      TcpListener tcpListener = new TcpListener(IPAddress.Any, 1337);
      tcpListener.Start();
      Console.WriteLine("Listening on port 1337");
      Console.WriteLine("Booting up.");
      while (true)
      {
        try
        {
          // ISSUE: variable of a compiler-generated type
          QBSessionManager instance = (QBSessionManager) Activator.CreateInstance(Type.GetTypeFromCLSID(new Guid("22E885D7-FB0B-49E3-B905-CCA6BD526B52")));
          // ISSUE: reference to a compiler-generated method
          instance.OpenConnection("", "TLS2");
          // ISSUE: reference to a compiler-generated method
          instance.BeginSession("G:\\Quickbooks\\GappData\\GappExpressV2\\GAPP Express Inc. 2015.qbw", ENOpenMode.omDontCare);
          Thread.Sleep(5000);
          Console.WriteLine("QuickBooks file opened.");
          Thread.Sleep(5000);
          Program.connection.Open();
          SqlDataReader sqlDataReader1 = new SqlCommand("SELECT * FROM Quickbooks WHERE type='AR'", Program.connection).ExecuteReader();
          while (sqlDataReader1.Read())
          {
            string str1 = sqlDataReader1["id"].ToString();
            string str2 = sqlDataReader1["type"].ToString();
            string str3 = sqlDataReader1["currency"].ToString();
            string str4 = sqlDataReader1["waybill"].ToString();
            string str5 = sqlDataReader1["invoice"].ToString();
            string val = sqlDataReader1["company"].ToString();
            string str6 = sqlDataReader1["service"].ToString();
            string str7 = sqlDataReader1["subtotal"].ToString();
            string str8 = sqlDataReader1["taxrate"].ToString();
            string s = sqlDataReader1["date"].ToString();
            Console.WriteLine("Round " + str1);
            Console.WriteLine("AR");
            // ISSUE: reference to a compiler-generated method
            // ISSUE: reference to a compiler-generated method
            // ISSUE: variable of a compiler-generated type
            IMsgSetRequest request = !(str3 == "CDN") && !(str3 == "CAD") ? instance.CreateMsgSetRequest("US", (short) 13, (short) 0) : instance.CreateMsgSetRequest("CA", (short) 13, (short) 0);
            // ISSUE: reference to a compiler-generated method
            // ISSUE: variable of a compiler-generated type
            IInvoiceAdd invoiceAdd = request.AppendInvoiceAddRq();
            // ISSUE: reference to a compiler-generated method
            invoiceAdd.CustomerRef.FullName.SetValue(val);
            if (str3 == "CDN" || str3 == "CAD")
            {
              // ISSUE: reference to a compiler-generated method
              invoiceAdd.ARAccountRef.FullName.SetValue("1300 ACCOUNTS RECEIVABLE (CAD)");
            }
            else
            {
              // ISSUE: reference to a compiler-generated method
              invoiceAdd.ARAccountRef.FullName.SetValue("1350 ACCOUNTS RECEIVABLE (USD)");
            }
            // ISSUE: reference to a compiler-generated method
            invoiceAdd.RefNumber.SetValue(str5.ToString());
            // ISSUE: reference to a compiler-generated method
            invoiceAdd.PONumber.SetValue(str4.ToString());
            // ISSUE: reference to a compiler-generated method
            invoiceAdd.TxnDate.SetValue(DateTime.Parse(s));
            // ISSUE: reference to a compiler-generated method
            // ISSUE: variable of a compiler-generated type
            IORInvoiceLineAdd orInvoiceLineAdd = invoiceAdd.ORInvoiceLineAddList.Append();
            if (str6.Trim() == "EX")
            {
              // ISSUE: reference to a compiler-generated method
              orInvoiceLineAdd.InvoiceLineAdd.ItemRef.FullName.SetValue("EXPEDITED GROUND R");
              // ISSUE: reference to a compiler-generated method
              orInvoiceLineAdd.InvoiceLineAdd.ORRatePriceLevel.Rate.SetValue(Convert.ToDouble(str7));
            }
            else if (str6.Trim() == "ST")
            {
              // ISSUE: reference to a compiler-generated method
              orInvoiceLineAdd.InvoiceLineAdd.ItemRef.FullName.SetValue("STORAGE R");
              // ISSUE: reference to a compiler-generated method
              orInvoiceLineAdd.InvoiceLineAdd.ORRatePriceLevel.Rate.SetValue(Convert.ToDouble(str7));
            }
            else if (str6.Trim() == "DS")
            {
              // ISSUE: reference to a compiler-generated method
              orInvoiceLineAdd.InvoiceLineAdd.ItemRef.FullName.SetValue("DESTUFFING R");
              // ISSUE: reference to a compiler-generated method
              orInvoiceLineAdd.InvoiceLineAdd.ORRatePriceLevel.Rate.SetValue(Convert.ToDouble(str7));
            }
            else
            {
              // ISSUE: reference to a compiler-generated method
              orInvoiceLineAdd.InvoiceLineAdd.ItemRef.FullName.SetValue("EXPEDITED GROUND R");
              // ISSUE: reference to a compiler-generated method
              orInvoiceLineAdd.InvoiceLineAdd.ORRatePriceLevel.Rate.SetValue(Convert.ToDouble(str7));
            }
            if (str8.ToString() == "13.00")
            {
              // ISSUE: reference to a compiler-generated method
              orInvoiceLineAdd.InvoiceLineAdd.SalesTaxCodeRef.FullName.SetValue("H13 ");
            }
            else if (str8.ToString() == "5.00")
            {
              // ISSUE: reference to a compiler-generated method
              orInvoiceLineAdd.InvoiceLineAdd.SalesTaxCodeRef.FullName.SetValue("G5");
            }
            else if (str8.ToString() == "14.00")
            {
              // ISSUE: reference to a compiler-generated method
              orInvoiceLineAdd.InvoiceLineAdd.SalesTaxCodeRef.FullName.SetValue("H14");
            }
            else if (str8.ToString() == "15.00")
            {
              // ISSUE: reference to a compiler-generated method
              orInvoiceLineAdd.InvoiceLineAdd.SalesTaxCodeRef.FullName.SetValue("H15");
            }
            else
            {
              // ISSUE: reference to a compiler-generated method
              orInvoiceLineAdd.InvoiceLineAdd.SalesTaxCodeRef.FullName.SetValue("E");
            }
            // ISSUE: reference to a compiler-generated method
            // ISSUE: variable of a compiler-generated type
            IMsgSetResponse msgSetResponse = instance.DoRequests(request);
            // ISSUE: reference to a compiler-generated method
            string xmlString = msgSetResponse.ToXMLString();
            XmlDocument xmlDocument = new XmlDocument();
            string xml = xmlString;
            xmlDocument.LoadXml(xml);
            string str9 = (string) null;
            string str10 = (string) null;
            string str11;
            try
            {
              str11 = xmlDocument.SelectSingleNode("QBXML/QBXMLMsgsRs/InvoiceAddRs/@statusSeverity").Value;
              str9 = xmlDocument.SelectSingleNode("QBXML/QBXMLMsgsRs/InvoiceAddRs/@statusMessage").Value;
              str10 = xmlDocument.SelectSingleNode("QBXML/QBXMLMsgsRs/InvoiceAddRs/@statusCode").Value;
            }
            catch (Exception ex)
            {
              str11 = "Good";
            }
            Console.WriteLine(str11 + str9);
            Program.connection2.Open();
            SqlCommand sqlCommand1 = new SqlCommand("INSERT INTO Quickbooks_Error (type, waybill, invoice, code, status, error, raw) VALUES (@type, @waybill, @invoice, @code, @status, @error, @raw)", Program.connection2);
            sqlCommand1.Parameters.AddWithValue("@type", (object) str2);
            sqlCommand1.Parameters.AddWithValue("@waybill", (object) str4);
            sqlCommand1.Parameters.AddWithValue("@invoice", (object) str5);
            sqlCommand1.Parameters.AddWithValue("@code", (object) str10);
            sqlCommand1.Parameters.AddWithValue("@status", (object) str11);
            sqlCommand1.Parameters.AddWithValue("@error", (object) str9);
            sqlCommand1.Parameters.AddWithValue("@raw", (object) xmlString);
            sqlCommand1.ExecuteNonQuery();
            SqlCommand sqlCommand2 = new SqlCommand("DELETE FROM Quickbooks WHERE id=@id", Program.connection2);
            sqlCommand2.Parameters.AddWithValue("@id", (object) str1);
            sqlCommand2.ExecuteNonQuery();
            Program.connection2.Close();
          }
          sqlDataReader1.Close();
          SqlDataReader sqlDataReader2 = new SqlCommand("SELECT * FROM Quickbooks WHERE type='AP'", Program.connection).ExecuteReader();
          while (sqlDataReader2.Read())
          {
            string str1 = sqlDataReader2["id"].ToString();
            string str2 = sqlDataReader2["type"].ToString();
            string str3 = sqlDataReader2["currency"].ToString();
            string val1 = sqlDataReader2["waybill"].ToString();
            string val2 = sqlDataReader2["invoice"].ToString();
            string val3 = sqlDataReader2["company"].ToString();
            string str4 = sqlDataReader2["service"].ToString();
            string str5 = sqlDataReader2["subtotal"].ToString();
            string str6 = sqlDataReader2["taxrate"].ToString();
            string s = sqlDataReader2["date"].ToString();
            Console.WriteLine("Round " + str1);
            Console.WriteLine("AP");
            // ISSUE: reference to a compiler-generated method
            // ISSUE: variable of a compiler-generated type
            IMsgSetRequest msgSetRequest = instance.CreateMsgSetRequest("CA", (short) 13, (short) 0);
            // ISSUE: reference to a compiler-generated method
            // ISSUE: variable of a compiler-generated type
            IBillAdd billAdd = msgSetRequest.AppendBillAddRq();
            // ISSUE: reference to a compiler-generated method
            billAdd.VendorRef.FullName.SetValue(val3);
            if (str3 == "CDN" || str3 == "CAD")
            {
              // ISSUE: reference to a compiler-generated method
              billAdd.APAccountRef.FullName.SetValue("2000 ACCOUNTS PAYABLE");
            }
            else
            {
              // ISSUE: reference to a compiler-generated method
              billAdd.APAccountRef.FullName.SetValue("2050 ACCOUNTS PAYABLE (USD)");
            }
            // ISSUE: reference to a compiler-generated method
            billAdd.RefNumber.SetValue(val2);
            // ISSUE: reference to a compiler-generated method
            billAdd.Memo.SetValue(val1);
            // ISSUE: reference to a compiler-generated method
            billAdd.TxnDate.SetValue(DateTime.Parse(s));
            // ISSUE: reference to a compiler-generated method
            // ISSUE: variable of a compiler-generated type
            IExpenseLineAdd expenseLineAdd = billAdd.ExpenseLineAddList.Append();
            if (str4.Trim() == "EX")
            {
              // ISSUE: reference to a compiler-generated method
              expenseLineAdd.AccountRef.FullName.SetValue("5021 EXPEDITED GROUND C");
              // ISSUE: reference to a compiler-generated method
              expenseLineAdd.Amount.SetValue(Convert.ToDouble(str5));
            }
            else if (str4.Trim() == "ST")
            {
              // ISSUE: reference to a compiler-generated method
              expenseLineAdd.AccountRef.FullName.SetValue("5070 STORAGE CS");
              // ISSUE: reference to a compiler-generated method
              expenseLineAdd.Amount.SetValue(Convert.ToDouble(str5));
            }
            else if (str4.Trim() == "DS")
            {
              // ISSUE: reference to a compiler-generated method
              expenseLineAdd.AccountRef.FullName.SetValue("5071 DESTUFFING CS");
              // ISSUE: reference to a compiler-generated method
              expenseLineAdd.Amount.SetValue(Convert.ToDouble(str5));
            }
            else
            {
              // ISSUE: reference to a compiler-generated method
              expenseLineAdd.AccountRef.FullName.SetValue("5021 EXPEDITED GROUND C");
              // ISSUE: reference to a compiler-generated method
              expenseLineAdd.Amount.SetValue(Convert.ToDouble(str5));
            }
            if (str6.ToString() == "13.00")
            {
              // ISSUE: reference to a compiler-generated method
              expenseLineAdd.SalesTaxCodeRef.FullName.SetValue("H13");
            }
            else if (str6.ToString() == "5.00")
            {
              // ISSUE: reference to a compiler-generated method
              expenseLineAdd.SalesTaxCodeRef.FullName.SetValue("G5");
            }
            else if (str6.ToString() == "14.00")
            {
              // ISSUE: reference to a compiler-generated method
              expenseLineAdd.SalesTaxCodeRef.FullName.SetValue("H14");
            }
            else if (str6.ToString() == "15.00")
            {
              // ISSUE: reference to a compiler-generated method
              expenseLineAdd.SalesTaxCodeRef.FullName.SetValue("H15");
            }
            else
            {
              // ISSUE: reference to a compiler-generated method
              expenseLineAdd.SalesTaxCodeRef.FullName.SetValue("E");
            }
            // ISSUE: reference to a compiler-generated method
            // ISSUE: variable of a compiler-generated type
            IMsgSetResponse msgSetResponse = instance.DoRequests(msgSetRequest);
            // ISSUE: reference to a compiler-generated method
            string xmlString = msgSetResponse.ToXMLString();
            XmlDocument xmlDocument = new XmlDocument();
            string xml = xmlString;
            xmlDocument.LoadXml(xml);
            string str7 = (string) null;
            string str8 = (string) null;
            string str9;
            try
            {
              str9 = xmlDocument.SelectSingleNode("QBXML/QBXMLMsgsRs/BillAddRs/@statusSeverity").Value;
              str7 = xmlDocument.SelectSingleNode("QBXML/QBXMLMsgsRs/BillAddRs/@statusMessage").Value;
              str8 = xmlDocument.SelectSingleNode("QBXML/QBXMLMsgsRs/BillAddRs/@statusCode").Value;
            }
            catch (Exception ex)
            {
              str9 = "Good";
            }
            Console.WriteLine(str9 + str7);
            Program.connection2.Open();
            SqlCommand sqlCommand1 = new SqlCommand("INSERT INTO Quickbooks_Error (type, waybill, invoice, code, status, error, raw) VALUES (@type, @waybill, @invoice, @code, @status, @error, @raw)", Program.connection2);
            sqlCommand1.Parameters.AddWithValue("@type", (object) str2);
            sqlCommand1.Parameters.AddWithValue("@waybill", (object) val1);
            sqlCommand1.Parameters.AddWithValue("@invoice", (object) val2);
            sqlCommand1.Parameters.AddWithValue("@code", (object) str8);
            sqlCommand1.Parameters.AddWithValue("@status", (object) str9);
            sqlCommand1.Parameters.AddWithValue("@error", (object) str7);
            sqlCommand1.Parameters.AddWithValue("@raw", (object) xmlString);
            sqlCommand1.ExecuteNonQuery();
            SqlCommand sqlCommand2 = new SqlCommand("DELETE FROM Quickbooks WHERE id=@id", Program.connection2);
            sqlCommand2.Parameters.AddWithValue("@id", (object) str1);
            sqlCommand2.ExecuteNonQuery();
            Program.connection2.Close();
          }
          sqlDataReader2.Close();
          Program.connection.Close();
        }
        catch (Exception ex)
        {
          tcpListener.Stop();
          Console.WriteLine(ex.ToString());
        }
        Console.WriteLine("Round Done, waiting 1hr");
        Thread.Sleep(300000);
      }
    }
  }
}
