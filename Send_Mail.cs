using System.Activities;
using System.ComponentModel;
using Microsoft.Office.Interop.Outlook;

namespace Kali_Mail
{
    public class Send_Mail : CodeActivity
    {
        [Category("Input")]
        [RequiredArgument]
        [Description("Enter recipient mail")]
        [DisplayName("To")]
        public InArgument<string> To { get; set; }

        [Category("Input")]
        [RequiredArgument]
        [Description("Enter mailbody")]
        [DisplayName("Body")]
        public InArgument<string> Body { get; set; }

        [Category("Input")]
        [RequiredArgument]
        [Description("Enter subject")]
        [DisplayName("Subject")]
        public InArgument<string> Subject { get; set; }

        [Category("Input")]
        [Description("Is the body HTML?")]
        [DisplayName("Is HTML")]
        public InArgument<bool> IsHtml { get; set; }

        [Category("Input")]
        [Description("Do not forward")]
        [DisplayName("Do not forward")]
        public InArgument<bool> doNotFoward { get; set; }

        protected override void Execute(CodeActivityContext context)
        {
            string recipient = To.Get(context);
            string mailBody = Body.Get(context);
            string mailSubject = Subject.Get(context);
            bool isHtml = IsHtml.Get(context);
            bool dnf = doNotFoward.Get(context);

            // Create a new Outlook application.
            Application outlookApp = new Application();

            // Create a new MailItem.
            MailItem mailItem = outlookApp.CreateItem(OlItemType.olMailItem) as MailItem;

            // Set recipient, subject, and body.
            mailItem.To = recipient;
            mailItem.Subject = mailSubject;
            
            if (isHtml)
                mailItem.BodyFormat = OlBodyFormat.olFormatHTML;
            else
                mailItem.BodyFormat = OlBodyFormat.olFormatPlain;
            
            // Set the permission to do not forward.
            if (dnf)
                mailItem.Permission = OlPermission.olDoNotForward;
            

            // Send the email.
            mailItem.Send();

            // Clean up.
            System.Runtime.InteropServices.Marshal.ReleaseComObject(mailItem);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(outlookApp);


        }
    }
}
