using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Olsens.Plugins.Common;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Olsens.Plugins.Children
{
    public class PostUpdateWrapper : PluginHelper
    {
        public PostUpdateWrapper(string unsecConfig, string securestring) : base(unsecConfig, securestring) { }

        protected override void Execute()
        {
            try
            {
                if (Context.MessageName.ToLower() != "update" || !Context.InputParameters.Contains("Target") || !(Context.InputParameters["Target"] is Entity)) return;

                AppendLog("Children PostUpdate - Plugin Excecution is Started.");

                Entity target = (Entity)Context.InputParameters["Target"];
                Entity postImage = Context.PostEntityImages.Contains("PostImage") ? (Entity)Context.PostEntityImages["PostImage"] : null;

                if ((target == null || target.LogicalName != "ols_children" || target.Id == Guid.Empty) || (postImage == null || postImage.Id == Guid.Empty))
                {
                    AppendLog("Target/PreImage/PostImage is null");
                    return;
                }

                #region On DOB Change
                if (target.Contains("ols_dateofbirth") || target.Contains("ols_lifestatus"))
                {
                    AppendLog("Target contains ols_dateofbirth");
                    int lifeStatus = postImage.Contains("ols_lifestatus") ? postImage.GetAttributeValue<OptionSetValue>("ols_lifestatus").Value : 0;
                    if (lifeStatus == Convert.ToInt32(Constants.ChildLifeStatus.Alive) || target.Contains("ols_firstgivenname"))
                    {
                        Guid funeralId = postImage.Contains("ols_funeralid") ? postImage.GetAttributeValue<EntityReference>("ols_funeralid").Id : Guid.Empty;
                        if (funeralId != Guid.Empty)
                        {
                            Entity funeral = Retrieve(UserType.User, "opportunity", funeralId, new ColumnSet("ols_dateofdeath", "ols_dateofdeathfrom"));
                            if (funeral != null)
                            {
                                DateTime? dod = funeral.Contains("ols_dateofdeath") ? funeral.GetAttributeValue<DateTime>("ols_dateofdeath") : (DateTime?)null;
                                if (dod == null)
                                    dod = funeral.Contains("ols_dateofdeathfrom") ? funeral.GetAttributeValue<DateTime>("ols_dateofdeathfrom") : (DateTime?)null;

                                if (dod != null)
                                {
                                    if (postImage.Contains("ols_dateofbirth"))
                                    {
                                        Age age = CalcAge(postImage.GetAttributeValue<DateTime>("ols_dateofbirth"), (DateTime)dod);

                                        Entity updateChild = new Entity(target.LogicalName, target.Id);
                                        updateChild["ols_age"] = Convert.ToString(age.Value);
                                        updateChild["ols_ageunit"] = new OptionSetValue(age.UnitValue);
                                        Update(UserType.User, updateChild);
                                    }
                                }
                            }
                        }
                    }
                    else if (lifeStatus == Convert.ToInt32(Constants.ChildLifeStatus.StillBorn))
                    {
                        Entity updateChild = new Entity(target.LogicalName, target.Id);
                        updateChild["ols_age"] = "0";
                        updateChild["ols_ageunit"] = new OptionSetValue(3);
                        Update(UserType.User, updateChild);
                    }
                    else
                    {
                        Entity updateChild = new Entity(target.LogicalName, target.Id);
                        updateChild["ols_age"] = null;
                        updateChild["ols_ageunit"] = null;
                        Update(UserType.User, updateChild);
                    }
                }
                else
                    AppendLog("Target doesn't contain ols_dateofbirth");
                #endregion

                #region On First given Name Change
                if (target.Contains("ols_firstgivenname"))
                {
                    Entity updateChild = new Entity(target.LogicalName, target.Id);
                    updateChild["ols_lifestatus"] = new OptionSetValue(1);
                    Update(UserType.User, updateChild);
                }
                #endregion

                #region On LifeStatus Change

                #endregion

                AppendLog("Children PostUpdate - Plugin Excecution is Completed.");

            }
            catch (Exception ex)
            {
                AppendLog("Error occured in Execute: " + ex.Message);
                throw new InvalidPluginExecutionException(ex.Message);
            }
        }

        public Age CalcAge(DateTime startDate, DateTime EndDate)
        {

            var dateDiff = new DateDiff(startDate, EndDate);
            var years = dateDiff.inYears;

            if (years < 0) return new Age(0, "years");
            var months = dateDiff.inMonths;

            var weeks = dateDiff.inWeeks;

            var days = dateDiff.inDays;

            if (years >= 1)
                return new Age(years, "years");
            else
               if (months >= 1) return new Age(months, "months");

            else
           if (weeks >= 1)
                return new Age(weeks, "weeks");
            if (days >= 1) return new Age(days, "days");
            else return new Age(0, "hours");
        }
    }
    public class Age
    {
        private int _value;
        private string _unit;
        private int _unitvalue;

        public int Value
        {
            get { return _value; }
            set { _value = value; }
        }
        public string Unit
        {
            get { return _unit; }
            set { _unit = value; }
        }
        public int UnitValue
        {
            get { return _unitvalue; }
            set { _unitvalue = value; }
        }
        public Age(int number, string unit)
        {
            _value = number;
            _unit = unit;
            switch (unit)
            {
                case "years":
                    _unitvalue = 6;
                    break;
                case "months":
                    _unitvalue = 4;
                    break;
                case "weeks":
                    _unitvalue = 5;
                    break;
                case "days":
                    _unitvalue = 1;
                    break;
                case "hours":
                    _unitvalue = 2;
                    break;
            }
        }

    }
    public class DateDiff
    {
        public int inYears
        {
            get; set;
        }
        public int inDays
        {
            get; set;
        }
        public int inWeeks
        {
            get; set;
        }
        public int inMonths
        {
            get; set;
        }

        public DateDiff(DateTime d1, DateTime d2)
        {
            TimeSpan span = d2 - d1;
            inDays = span.Days;//Convert.ToInt32((d2 - d1)) / (1000 * 60 * 60 * 24);
            inWeeks = inDays / 7;// Convert.ToInt32(Math.Floor(Convert.ToDecimal(Convert.ToInt32(d2 - d1) / 1000 / 60 / 60 / 24 / 7)));

            var d1Y = d1.Year;
            var d2Y = d2.Year;
            var d1M = d1.Month;
            var d2M = d2.Month;
            var d2d = d2.Date;
            var d1d = d1.Date;
            if ((d2M + 12 * d2Y) - (d1M + 12 * d1Y) > 1 || (d2d > d1d))
                inMonths = (d2M + 12 * d2Y) - (d1M + 12 * d1Y);
            else if (((d2M + 12 * d2Y) - (d1M + 12 * d1Y) == 1) && (d2d == d1d)) inMonths = 1;
            else inMonths = 0;

            var m1 = d1.Month;
            var m2 = d2.Month;
            var y1 = d1.Year;
            var y2 = d2.Year;
            var dd1 = d1.Date;
            var dd2 = d2.Date;
            if (y2 > y1)
            {
                if (m2 > m1) inYears = y2 - y1;
                else
                    if (m2 < m1) inYears = y2 - y1 - 1;
                else
                        if (m2 == m1)
                {
                    if (dd2 >= dd1) inYears = y2 - y1;
                    else inYears = y2 - y1 - 1;
                }
            }
            else if (y1 > y2) inYears = -1;
            else inYears = 0;
        }
    }
    public class PostUpdate : IPlugin
    {
        string UnsecConfig = string.Empty;
        string Securestring = string.Empty;
        public PostUpdate(string unsecConfig, string securestring)
        {
            UnsecConfig = unsecConfig;
            Securestring = securestring;
        }
        public void Execute(IServiceProvider serviceProvider)
        {
            var pluginCode = new PostUpdateWrapper(UnsecConfig, Securestring);
            pluginCode.Execute(serviceProvider);
            pluginCode.Dispose();
        }
    }
}
