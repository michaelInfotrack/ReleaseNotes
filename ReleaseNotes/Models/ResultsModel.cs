using Atlassian.Jira;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Web;

namespace ReleaseNotes.Models
{
    public class ResultsModel
    {
        public List<Issue> JiraIssues { get; set; }

        public enum ProjectTypes
        {
            [Description("Development Management Team")]
            DMT,
            [Description("Global Platform")]
            GLOB,
            [Description("VOI")]
            VOI,
            [Description("iMajor")]
            IMAJOR,
            [Description("Infotrack UK")]
            UK,
            [Description("Internal")]
            IN,
            [Description("LABS")]
            LABS,
            [Description("MapIT")]
            MAPIT,
            [Description("Maple")]
            MAP,
            [Description("Pencil")]
            PEN,
            [Description("Pexa")]
            PEXA,
            [Description("PlanIT")]
            PLN,
            [Description("Reveal")]
            REV,
            [Description("SettleIT")]
            SC,
            [Description("SignIT")]
            SIG,
            [Description("Test")]
            TEST,
            [Description("TrackIT")]
            TIT,
            [Description("US - The List")]
            UL,
            [Description("US Platform")]
            USP,
            [Description("Website")]
            WEB,
            [Description("We Care")]
            CAR,
            [Description("Other")]
            Default
        }
    }
}