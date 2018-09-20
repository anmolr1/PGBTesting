using Microsoft.VisualStudio.TestTools.UITest.Extension;
using Microsoft.VisualStudio.TestTools.UITesting;
using Microsoft.VisualStudio.TestTools.UITesting.HtmlControls;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PGBtesting
{
    public class RegisterForm
    {
        private BrowserWindow mbrowserWindow;

        /* This method Registration Form Locate the Excel file and take data from the file.
            Identifies the controls of the form and insert data from excel to the form input fields. */

        public void RegistrationForm()
        {
            int x = 7;
            ExceltToForm util = new ExceltToForm(); 
            util.PopulateInCollection(@"D:\Coded ui\Emp_data1.xlsx");
       
            EnterText(new HtmlComboBox(), PropertyType.Id, "SubGroupSubDivisionId", util.ReadData(x, "SubGroup Division"));
            EnterText(new HtmlComboBox(), PropertyType.Id, "BenefitClassDropDown", util.ReadData(x, "Benefit Class"));
            EnterText(new HtmlComboBox(), PropertyType.Id, "EnrollmentTypeId", util.ReadData(x, "ENROLLMENT_TYPE"));
            EnterText(new HtmlComboBox(), PropertyType.Id, "EmploymentType", "Contract");

            EnterText(new HtmlEdit(), PropertyType.Id, "EmployeeId", util.ReadData(x, "EMPLOYEE_ID"));
            EnterText(new HtmlEdit(), PropertyType.Id, "FirstName", util.ReadData(x, "FIRST_NAME"));
            EnterText(new HtmlEdit(), PropertyType.Id, "LastName", util.ReadData(x, "LAST_NAME"));

            EnterText(new HtmlComboBox(), PropertyType.Id, "ImmegrationStatus", util.ReadData(x, "IMMIGRATIONSTATUS"));

            EnterText(new HtmlEdit(), PropertyType.Id, "AddressLine1", util.ReadData(x, "EMP_ADDRESS1"));

            EnterText(new HtmlComboBox(), PropertyType.Id, "CountryId", util.ReadData(x, "COUNTRY"));

            EnterText(new HtmlEdit(), PropertyType.Id, "City", util.ReadData(x, "EMP_CITY"));
            EnterText(new HtmlEdit(), PropertyType.Id, "Zip", util.ReadData(x, "EMP_ZIP"));
            EnterText(new HtmlEdit(), PropertyType.Id, "StateName", util.ReadData(x, "EMP_STATE"));

            EnterText(new HtmlComboBox(), PropertyType.Id, "Gender", util.ReadData(x, "GENDER"));
            EnterText(new HtmlComboBox(), PropertyType.Id, "MaritalStatus", util.ReadData(x, "MARITAL_STATUS"));

            EnterText(new HtmlEdit(), PropertyType.Id, "DateOfBirth", util.ReadData(x, "DATE_OF_BIRTH"));
            EnterText(new HtmlEdit(), PropertyType.Id, "DateOfHire", util.ReadData(x, "DATE_OF_HIRE"));
            EnterText(new HtmlEdit(), PropertyType.Id, "EmergencyContactPerson", util.ReadData(x, "Emergency Contact"));
            EnterText(new HtmlEdit(), PropertyType.Id, "txtUser", util.ReadData(x, "Username"));
            EnterText(new HtmlEdit(), PropertyType.Id, "BaseSalary", util.ReadData(x, "ANNUAL_EARNINGS"));

            mouseClick(new HtmlInputButton(), PropertyType.DisplayText, "Save");
        }

        public enum PropertyType
        {
            Id,
            Name,
            ClassName,
            InnerText,
            TagInstance,
            DisplayText
        }

        // This is a generic function to handle all types of mouse click actions.
        public void mouseClick(object controlType, PropertyType type, string propertyValue)
        {
            if (controlType is HtmlHyperlink)
            {
                HtmlHyperlink link = new HtmlHyperlink(ParentWindow);
                if (type == PropertyType.Name)
                {
                    link.SearchProperties[HtmlHyperlink.PropertyNames.Name] = propertyValue;
                }
                else if (type == PropertyType.InnerText)
                {
                    link.SearchProperties[HtmlHyperlink.PropertyNames.InnerText] = propertyValue;
                }
                else if (type == PropertyType.Id)
                {
                    link.SearchProperties[HtmlHyperlink.PropertyNames.Id] = propertyValue;
                }
                Mouse.Click(link);
            }
            else if (controlType is HtmlButton)
            {
                HtmlButton btn = new HtmlButton(ParentWindow);
                if (type == PropertyType.Name)
                {
                    btn.SearchProperties[HtmlButton.PropertyNames.Name] = propertyValue;
                }
                else if (type == PropertyType.DisplayText)
                {
                    btn.SearchProperties[HtmlButton.PropertyNames.DisplayText] = propertyValue;
                }
                else if (type == PropertyType.Id)
                {
                    btn.SearchProperties[HtmlButton.PropertyNames.Id] = propertyValue;
                }
                Mouse.Click(btn);
            }
            else if (controlType is HtmlInputButton)
            {
                HtmlInputButton btn = new HtmlInputButton(ParentWindow);
                if (type == PropertyType.Name)
                {
                    btn.SearchProperties[HtmlInputButton.PropertyNames.Name] = propertyValue;
                }
                else if (type == PropertyType.DisplayText)
                {
                    btn.SearchProperties[HtmlButton.PropertyNames.DisplayText] = propertyValue;
                }
                else if (type == PropertyType.Id)
                {
                    btn.SearchProperties[HtmlInputButton.PropertyNames.Id] = propertyValue;
                }
                Mouse.Click(btn);
            }

        }

        // This function controls all kinds of input controls.

        public void EnterText(object controlType, PropertyType type, string propertyValue, string txt)
        {
            if (controlType is HtmlEdit)
            {
                HtmlEdit edit = new HtmlEdit(ParentWindow);
                if (type == PropertyType.Name)
                {
                    edit.SearchProperties[HtmlEdit.PropertyNames.Name] = propertyValue;
                }
                else if (type == PropertyType.Id)
                {
                    edit.SearchProperties[HtmlEdit.PropertyNames.Id] = propertyValue;
                }
                else if (type == PropertyType.ClassName)
                {
                    edit.SearchProperties[HtmlEdit.PropertyNames.ClassName] = propertyValue;
                }
                Keyboard.SendKeys(edit, txt);
            }
            else if (controlType is HtmlComboBox)
            {
                HtmlComboBox edit = new HtmlComboBox(ParentWindow);
                if (type == PropertyType.Name)
                {
                    edit.SearchProperties[HtmlComboBox.PropertyNames.Name] = propertyValue;
                }
                else if (type == PropertyType.Id)
                {
                    edit.SearchProperties[HtmlComboBox.PropertyNames.Id] = propertyValue;
                }
                else if (type == PropertyType.ClassName)
                {
                    edit.SearchProperties[HtmlComboBox.PropertyNames.ClassName] = propertyValue;
                }
                edit.SelectedItem = txt;
                Mouse.Click(edit);
            }
        }

        public BrowserWindow GetBrowserWindow()
        {
            BrowserWindow window = new BrowserWindow();
            window.SearchProperties[UITestControl.PropertyNames.ClassName] = BrowserWindow.CurrentBrowser.ToString();
            return window;
        }
        public BrowserWindow ParentWindow
        {
            get
            {
                if (this.mbrowserWindow == null)
                {
                    this.mbrowserWindow = GetBrowserWindow();
                }
                return this.mbrowserWindow;
            }
        }

    }
}
