using System;
using System.Globalization;
using System.Windows.Data;
using Microsoft.ActiveDirectory.Management;
using System.Collections.Generic;

    public class ADPropertyValueCollectionConverter : IValueConverter
    {

        public string Separator
        {
            get { return ";"; }
        }

        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (value != null)
            {
                string returnString = string.Empty;
                ADPropertyValueCollection collection = (ADPropertyValueCollection)value;
                if (collection.Count > 0)
                {
                    for(int i = 0; i < collection.Count; i++)
                    {
                        returnString = returnString + collection[i].ToString() + ";";
                    }
                }
                if(returnString.EndsWith(";"))
                {
                    returnString = returnString.Remove(returnString.Length - 1);
                }
                return returnString;
            }
            else
            {
                return null;
            }
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }

    public class FileTimeConverter : IValueConverter
    {
        private string dateTimeFormat;

        public FileTimeConverter(string _dateTimeFormat)
        {
            this.dateTimeFormat = _dateTimeFormat;
        }
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (value != null && (long)value > 0)
            {
                return DateTime.FromFileTime((long)value).ToString(dateTimeFormat);
            }
            else
            {
                return String.Empty;
            }
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }

    public class DateFormatConverter : IValueConverter
    {
        private string dateTimeFormat;

        public DateFormatConverter(string _dateTimeFormat)
        {
            this.dateTimeFormat = _dateTimeFormat;
        }
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (value != null)
            {
                DateTime dt = (DateTime)value;
                return dt.ToString(dateTimeFormat);
            }
            else
            {
                return null;
            }
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }

    public class msExchRemoteRecipientTypeConverter : IValueConverter
    {
        Dictionary<Int64, string> dictionary = new Dictionary<Int64, string>()
        {
            {1,   "ProvisionMailbox"},
            {2,   "ProvisionArchive (On-Prem Mailbox)"},
            {3,   "ProvisionMailbox, ProvisionArchive"},
            {4,   "Migrated (UserMailbox)"},
            {6,   "ProvisionArchive, Migrated"},
            {8,   "DeprovisionMailbox"},
            {10,  "ProvisionArchive, DeprovisionMailbox"},
            {16,  "DeprovisionArchive (On-Prem Mailbox)"},
            {17,  "ProvisionMailbox, DeprovisionArchive"},
            {20,  "Migrated, DeprovisionArchive"},
            {24,  "DeprovisionMailbox, DeprovisionArchive"},
            {33,  "ProvisionMailbox, RoomMailbox"},
            {35,  "ProvisionMailbox, ProvisionArchive, RoomMailbox"},
            {36,  "Migrated, RoomMailbox"},
            {38,  "ProvisionArchive, Migrated, RoomMailbox"},
            {49,  "ProvisionMailbox, DeprovisionArchive, RoomMailbox"},
            {52,  "Migrated, DeprovisionArchive, RoomMailbox"},
            {65,  "ProvisionMailbox, EquipmentMailbox"},
            {67,  "ProvisionMailbox, ProvisionArchive, EquipmentMailbox"},
            {68,  "Migrated, EquipmentMailbox"},
            {70,  "ProvisionArchive, Migrated, EquipmentMailbox"},
            {81,  "ProvisionMailbox, DeprovisionArchive, EquipmentMailbox"},
            {84,  "Migrated, DeprovisionArchive, EquipmentMailbox"},
            {100, "Migrated, SharedMailbox"},
            {102, "ProvisionArchive, Migrated, SharedMailbox"},
            {116, "Migrated, DeprovisionArchive, SharedMailbox"}
        };

        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (value is Int64)
            {
                if (dictionary.ContainsKey((Int64)value))
                {
                    return dictionary[(Int64)value];
                }
                else
                {
                    return value;
                }
            }
            else
            {
                return null;
            }
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }

    public class msExchRecipientDisplayTypeConverter : IValueConverter
    {
        Dictionary<int, string> dictionary = new Dictionary<int, string>()
        {
            {-2147483642,   "MailUser (RemoteUserMailbox)"},
            {-2147481850,   "MailUser (RemoteRoomMailbox)"},
            {-2147481594,   "MailUser (RemoteEquipmentMailbox)"},
            {0,             "UserMailbox (shared)"},
            {1,             "MailUniversalDistributionGroup"},
            {6,             "MailContact"},
            {7,             "UserMailbox (room)"},
            {8,             "UserMailbox (equipment)"},
            {1073741824 ,   "UserMailbox"},
            {1073741833 ,   "MailUniversalSecurityGroup"}
        };

        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (value is int)
            {
                if (dictionary.ContainsKey((int)value))
                {
                    return dictionary[(int)value];
                }
                else
                {
                    return value;
                }
            }
            else
            {
                return null;
            }
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }

    public class managedByConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (value is string)
            {
                string val = (string)value;
                if (val.Length > 4 && val.Contains(","))
                {
                    string retValue = (string)val;
                    retValue = retValue.Substring(3);
                    retValue = retValue.Substring(0, retValue.IndexOf(','));
                    return retValue;
                }
                else
                {
                    return null;
                }
            }
            else
            {
                return null;
            }
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }