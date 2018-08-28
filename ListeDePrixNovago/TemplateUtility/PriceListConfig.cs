using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Configuration;
using System.Globalization;
using System.ComponentModel;
using System.Collections.Specialized;
using System.Xml.Serialization;
using System.IO;
using System.Security.Cryptography;

namespace ListeDePrixNovago.PDFTemplate
{
    public sealed class PriceListConfig
    {
        private string logoPath;
        private string title;
        private string footer;
        private bool isValidityDateInFooter;
        private string smtpServer;
        private int smtpPort;
        private string smtpUsername;
        private string smtpPassword;
        private string teamsGroupId;
        private string driveItemId;

        public string LogoPath { get => logoPath; set => logoPath = value; }
        public string Title { get => title; set => title = value; }
        public string Footer { get => footer; set => footer = value; }
        public bool IsValidityDateInFooter { get => isValidityDateInFooter; set => isValidityDateInFooter = value; }
        public string SmtpServer { get => smtpServer; set => smtpServer = value; }
        public string SmtpUsername { get => smtpUsername; set => smtpUsername = value; }
        public string SmtpPassword { get => smtpPassword; set => smtpPassword = value; }
        public int SmtpPort { get => smtpPort; set => smtpPort = value; }
        public string TeamsGroupId { get => teamsGroupId; set => teamsGroupId = value; }
        public string DriveItemId { get => driveItemId; set => driveItemId = value; }
    }
}
