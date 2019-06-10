using System;
using System.Collections.Generic;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Microsoft.Win32;
using mshtml;
using SHDocVw;

namespace IEExtension
{
    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.None)]
    [Guid("D40C654D-7C51-4EB3-95B2-1E23905C2A2D")]
    [ProgId("MyBHO.1950LabsPOC")]
    public class POC1950LabsBHO : IObjectWithSite, IOleCommandTarget
    {
        IWebBrowser2 browser;
        private object site;

        #region Highlight Text
        void OnDocumentComplete(object pDisp, ref object URL)
        {
            try
            {
                var userName = Environment.UserName; 
                var document2 = browser.Document as IHTMLDocument2;
                var document3 = browser.Document as IHTMLDocument3;

                var window = document2.parentWindow;
                window.execScript("function pluginInit() { " +

                    " var otherPages = 'www.booking.com www.airbnb.com www.latam.com www.copaair.com www.es.kayak.com www.en.kayak.com www.kayak.com';" +
                    "var checkoutFlag = false; var flightSetup = false; var flightConfirm = false; var flightSummary = false; var flightFinalStep = false; " +
                    "var fromAirport = ''; var toAirport = ''; var fromDate = ''; var toDate = ''; var totalCost = '';" +
                    "if((window.location.hostname === 'www.aa.com'|| otherPages.indexOf(window.location.hostname) >= 0) && document.getElementById('main1950button') === null) { " +
                        "var div = document.getElementsByTagName('body')[0]; " +
                        "var newElement = document.createElement('div');" +
                        "newElement.innerHTML = \"" +
                        "<div id='main1950button' style='width: 100px; height: 100px; right: 50px; bottom: 50px; position: fixed; z-index: 9999; cursor: pointer;'>" +
                            "<img style='width: 100px; height: 100px;' src='https://firebasestorage.googleapis.com/v0/b/labs-51e50.appspot.com/o/logo.png?alt=media&amp;token=c80334a0-2a40-48d7-87f5-32f83a3eca8f' />" +
                        "</div>" +
                        "<div id='mainWidgetPanel' style='color: white; border-radius: 5px; visibility: hidden; z-index:9999; width: 300px; background-color: black; height: 500px; position: fixed; display: block; top:40px; right: 40px;'>" +
                            "<div style=' text-align: right; background-color: #f1f1f1; border: 2px solid black; border-top-left-radius: 5px; border-top-right-radius: 5px; height: 40px;'><img style='width: 20px; margin: 8px;' src='https://static.thenounproject.com/png/7294-200.png' id='mainWidgetPanelClose' /></div>" +
                            "<img style='width: 100px;  margin: auto; margin-top:20px; display: block' src='https://firebasestorage.googleapis.com/v0/b/labs-51e50.appspot.com/o/logo.png?alt=media&amp;token=c80334a0-2a40-48d7-87f5-32f83a3eca8f' />" +
                            "<div style='width: 250px; margin: auto; margin-top: 15px;'>" +
                                "<div style='padding: 5px;' id='userName'><strong>USER: </strong>" + userName + "</div>" +
                                "<div style='padding: 5px;' id='pageUrl'><strong>URL: </strong>\" + window.location.origin + \"</div>" +

                                "<div style='padding: 5px;'>" +
                                    "<label id='step'></label>" +
                                "</div>" +
                                "<div style='padding: 5px;'>" +
                                    "<label id='fromAirport'></label>" +
                                "</div>" +
                                "<div style='padding: 5px;'>" +
                                    "<label id='toAirport'></label>" +
                                "</div>" +
                                "<div id='totalCostDiv' style='padding: 5px;'>" +
                                    "<label id='totalCost'></label>" +
                                "</div>" +
                                "<button id='checkoutBtn' style='width:100%; color: white;  display: none; cursor: pointer;border:none;margin-top: 10px;text-align: center;margin-left: auto;margin-right: auto;padding-top: 7px;padding-bottom: 7px; background-color: #578db9;'>Checkout</button>" +
                            "</div>" +
                        "</div>\"; " +
                        "div.appendChild(newElement);" +
                        "document.getElementById('main1950button').addEventListener('click', function() { document.getElementById('mainWidgetPanel').style.visibility = 'visible';});" +
                        "document.getElementById('mainWidgetPanelClose').addEventListener('click', function() { document.getElementById('mainWidgetPanel').style.visibility = 'hidden';});" +
                        "document.getElementById('checkoutBtn').addEventListener('click', function() { document.getElementById('button_continue_guest').click();});" +
                    "}" +

                    " if(window.location.href.indexOf('aa.com/booking/flights/choose-flights/flight1') > 0 && !flightConfirm) { " +
                    "flightConfirm = true;" +
                     "fromAirport = document.querySelector('#slice0Flight1 > div > div:nth-child(1) > div > div:nth-child(1) > div.span4.span-phone6 > span').innerText;" +
                     "toAirport = document.querySelector('#slice0Flight1 > div > div:nth-child(1) > div > div:nth-child(1) > div.span4.span-phone5 > span').innerText;" +
                     "document.getElementById('fromAirport').innerHTML = '<strong>ORIGIN: </strong>' + fromAirport; " +
                     "document.getElementById('toAirport').innerHTML = '<strong>DESTINATION: </strong>' + toAirport; " +
                     "document.getElementById('step').innerHTML = '<strong>STEP: </strong>2 of 5';" +
                    "}" +

                    " if(window.location.href.indexOf('aa.com/booking/flights/choose-flights/flight2') > 0 && !flightConfirm) { " +
                    "flightConfirm = true;" +
                     "fromAirport = document.querySelector('#slice1Flight1 > div > div:nth-child(1) > div > div:nth-child(1) > div.span4.span-phone6 > span').innerText;" +
                     "toAirport = document.querySelector('#slice1Flight1 > div > div:nth-child(1) > div > div:nth-child(1) > div.span4.span-phone5 > span').innerText;" +
                     "document.getElementById('fromAirport').innerHTML = '<strong>ORIGIN: </strong>' + fromAirport; " +
                     "document.getElementById('toAirport').innerHTML = '<strong>DESTINATION: </strong>' + toAirport; " +
                     "document.getElementById('step').innerHTML = '<strong>STEP: </strong>3 of 5';" +
                    "}" +

                    " if(window.location.href.indexOf('our-trip-summary?selectedFareId') > 0 && !flightSummary) { " +
                    "flightSummary = true;" +
                     "fromAirport = document.querySelector('#tripSummaryForm > section:nth-child(3) > div > div > div.span8 > div > div:nth-child(1) > div.flight-summary > h2 > span:nth-child(2)').firstChild.innerHTML;" +
                     "toAirport = document.querySelector('#tripSummaryForm > section:nth-child(3) > div > div > div.span8 > div > div:nth-child(1) > div.flight-summary > h2 > span:nth-child(4)').firstChild.innerHTML;" +
                     "totalCost = document.querySelector('#tripSummaryForm > section:nth-child(3) > div > div > div.span4 > div > div.unit-price > span.cost').innerText;" +
                     "document.getElementById('fromAirport').innerHTML = '<strong>ORIGIN: </strong>' + fromAirport; " +
                     "document.getElementById('toAirport').innerHTML = '<strong>DESTINATION: </strong>' + toAirport; " +
                     "document.getElementById('totalCost').innerHTML = '<strong>TOTAL: </strong>' + totalCost; " +
                     "document.getElementById('totalCostDiv').style.border = '1px solid white';" +
                     "document.getElementById('checkoutBtn').style.display = 'block';" +
                     "document.getElementById('step').innerHTML = '<strong>STEP: </strong>4 of 5';" +
                    "}" +

                    " if(window.location.href.indexOf('aa.com/booking/passengers?bookingPathStateId') > 0 && !flightFinalStep) { " +
                    "flightFinalStep = true;" +
                     "document.getElementById('totalCostDiv').style.border = 'none';" +
                     "document.getElementById('checkoutBtn').style.display = 'none';" +
                     "document.getElementById('step').innerHTML = '<strong>STEP: </strong>5 of 5';" +
                    "}" +

                    " if(window.location.href.indexOf('aa.com/homePage.do') > 0 && !flightSetup) { " +
                    "flightSetup = true;" +

                        "document.getElementById('reservationFlightSearchForm.originAirport').addEventListener('input', function(evt) { " +
                            "fromAirport = document.getElementById('reservationFlightSearchForm.originAirport').value;" +
                            "document.getElementById('fromAirport').innerHTML = fromAirport; " +

                        "});" +

                        "document.getElementById('reservationFlightSearchForm.destinationAirport').addEventListener('input', function(evt) { " +
                            "toAirport = document.getElementById('reservationFlightSearchForm.destinationAirport').value;" +
                            "document.getElementById('toAirport').innerHTML = toAirport; " +

                        "});" +

                        "toAirport = document.getElementById('reservationFlightSearchForm.destinationAirport').value;" +
                        "document.getElementById('step').innerHTML = '<strong>STEP: </strong>1 of 5';" +
                        "document.getElementById('fromAirport').innerHTML = document.getElementById('reservationFlightSearchForm.originAirport').value;" + 


                    "document.getElementById('toAirport').innerHTML = toAirport; " +
                    "}" +
                    "} window.onload = pluginInit ;"
                    );
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
        #endregion

        [DllImport("ieframe.dll")]
        public static extern int IEGetWriteableHKCU(ref IntPtr phKey);

        [Guid("6D5140C1-7436-11CE-8034-00AA006009FA")]
        [InterfaceType(1)]
        public interface IServiceProvider
        {
            int QueryService(ref Guid guidService, ref Guid riid, out IntPtr ppvObject);
        }

        #region Implementation of IObjectWithSite
        int IObjectWithSite.SetSite(object site)
        {
            this.site = site;

            if (site != null)
            {
                var serviceProv = (IServiceProvider)this.site;
                var guidIWebBrowserApp = Marshal.GenerateGuidForType(typeof(IWebBrowserApp)); 
                var guidIWebBrowser2 = Marshal.GenerateGuidForType(typeof(IWebBrowser2));
                IntPtr intPtr;
                serviceProv.QueryService(ref guidIWebBrowserApp, ref guidIWebBrowser2, out intPtr);

                browser = (IWebBrowser2)Marshal.GetObjectForIUnknown(intPtr);

                ((DWebBrowserEvents2_Event)browser).DocumentComplete +=
                    new DWebBrowserEvents2_DocumentCompleteEventHandler(this.OnDocumentComplete);
            }
            else
            {
                ((DWebBrowserEvents2_Event)browser).DocumentComplete -=
                    new DWebBrowserEvents2_DocumentCompleteEventHandler(this.OnDocumentComplete);
                browser = null;
            }
            return 0;
        }
        int IObjectWithSite.GetSite(ref Guid guid, out IntPtr ppvSite)
        {
            IntPtr punk = Marshal.GetIUnknownForObject(browser);
            int hr = Marshal.QueryInterface(punk, ref guid, out ppvSite);
            Marshal.Release(punk);
            return hr;
        }
        #endregion
        #region Implementation of IOleCommandTarget
        int IOleCommandTarget.QueryStatus(IntPtr pguidCmdGroup, uint cCmds, ref OLECMD prgCmds, IntPtr pCmdText)
        {
            return 0;
        }
        int IOleCommandTarget.Exec(IntPtr pguidCmdGroup, uint nCmdID, uint nCmdexecopt, IntPtr pvaIn, IntPtr pvaOut)
        {
            try
            {
                var document = browser.Document as IHTMLDocument2;
                var window = document.parentWindow;
                var form = new IE_POC_Form("");
                form.StartPosition = FormStartPosition.Manual;
                form.Location = new Point(Cursor.Position.X, Cursor.Position.Y);
                form.ShowDialog();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

            return 0;
        }
        #endregion

        #region Registering with regasm
        public static string RegBHO = "Software\\Microsoft\\Windows\\CurrentVersion\\Explorer\\Browser Helper Objects";
        public static string RegCmd = "Software\\Microsoft\\Internet Explorer\\Extensions";

        [ComRegisterFunction]
        public static void RegisterBHO(Type type)
        {
            string guid = type.GUID.ToString("B");

            // BHO
            {
                RegistryKey registryKey = Registry.LocalMachine.OpenSubKey(RegBHO, true);
                if (registryKey == null)
                    registryKey = Registry.LocalMachine.CreateSubKey(RegBHO);
                RegistryKey key = registryKey.OpenSubKey(guid);
                if (key == null)
                    key = registryKey.CreateSubKey(guid);
                key.SetValue("Alright", 1);
                registryKey.Close();
                key.Close();
            }

            // Command
            {
                RegistryKey registryKey = Registry.LocalMachine.OpenSubKey(RegCmd, true);
                if (registryKey == null)
                    registryKey = Registry.LocalMachine.CreateSubKey(RegCmd);
                RegistryKey key = registryKey.OpenSubKey(guid);
                if (key == null)
                    key = registryKey.CreateSubKey(guid);
                key.SetValue("ButtonText", "POC IE Plugin");
                key.SetValue("CLSID", "{1FBA04EE-3024-11d2-8F1F-0000F87ABD16}");
                key.SetValue("ClsidExtension", guid);
                key.SetValue("Icon", @"C:\Users\Admin\Desktop\IE_Ext_test\1950Icon.ico");
                key.SetValue("HotIcon", @"C:\Users\Admin\Desktop\IE_Ext_test\1950Icon.ico");
                key.SetValue("Default Visible", "Yes");
                key.SetValue("MenuText", "&POC IE Plugin");
                key.SetValue("ToolTip", "POC IE Plugin");
                registryKey.Close();
                key.Close();
            }
        }

        [ComUnregisterFunction]
        public static void UnregisterBHO(Type type)
        {
            string guid = type.GUID.ToString("B");
            // BHO
            {
                RegistryKey registryKey = Registry.LocalMachine.OpenSubKey(RegBHO, true);
                if (registryKey != null)
                    registryKey.DeleteSubKey(guid, false);
            }
            // Command
            {
                RegistryKey registryKey = Registry.LocalMachine.OpenSubKey(RegCmd, true);
                if (registryKey != null)
                    registryKey.DeleteSubKey(guid, false);
            }
        }
        #endregion
    }
}