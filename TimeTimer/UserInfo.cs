using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;


namespace TimeTimer
{
    public class UserInfo
    {
        private string name;

        private TimeSpan periodWork;
        private TimeSpan periodBreakDinner;
        private TimeSpan periodBreakPause;
        private TimeSpan periodBreakMeeting;
        private TimeSpan periodBreakStudy;
        private TimeSpan periodBreakNote;
        private TimeSpan periodBreakDoctor;

        private string typeDinner;
        private string typePause;
        private string typeMeeting;
        private string typeStudy;
        private string typeNote;
        private string typeDoctor;
        private string typeWork;



        private string noteMeeting;
        private string noteStudy;

        public string Name { get { return name; } set { name = value; } }
        public string NoteMeeting { get { return noteMeeting; } set { noteMeeting = value; } }
        public string NoteStudy { get { return noteStudy; } set { noteStudy = value; } }
        

        public TimeSpan BreakDinner { get { return periodBreakDinner; } set { periodBreakDinner = value; } }
        public TimeSpan BreakPause { get { return periodBreakPause; } set { periodBreakPause = value; } }
        public TimeSpan BreakMeeting { get { return periodBreakMeeting; } set { periodBreakMeeting = value; } }
        public TimeSpan BreakStudy { get { return periodBreakStudy; } set { periodBreakStudy = value; } }
        public TimeSpan BreakNote { get { return periodBreakNote; } set { periodBreakNote = value; } }
        public TimeSpan BreakDoctor { get { return periodBreakDoctor; } set { periodBreakDoctor = value; } }
        public TimeSpan PeriodWork { get { return periodWork; } set { periodWork = value; } }

        public string TypeDinner { get { return typeDinner; } set { typeDinner = value; } }
        public string TypePause { get { return typePause; } set { typePause = value; } }
        public string TypeMeeting { get { return typeMeeting; } set { typeMeeting = value; } }
        public string TypeStudy { get { return typeStudy; } set { typeStudy = value; } }
        public string TypeNote { get { return typeNote; } set { typeNote = value; } }
        public string TypeDoctor { get { return typeDoctor; } set { typeDoctor = value; } }
        public string TypeWork { get { return typeWork; } set { typeWork = value; } }

 
        public double ProcentWorks(string num)
        {
            double workPeriods;
            double breakPeriods;
            workPeriods = Convert.ToDouble(PeriodWork.TotalMinutes);
            breakPeriods = Convert.ToDouble(BreakDinner.TotalMinutes) + Convert.ToDouble(BreakDoctor.TotalMinutes) + Convert.ToDouble(BreakMeeting.TotalMinutes) + Convert.ToDouble(BreakNote.TotalMinutes) + Convert.ToDouble(BreakPause.TotalMinutes) + Convert.ToDouble(BreakStudy.TotalMinutes);
            double procentWork = 0;
            double procentBreak = 0;

            procentWork = workPeriods;
            procentBreak = breakPeriods;

            procentBreak = Math.Round((procentBreak / procentWork) * 100);

            if (num == "w")
                return Math.Round(100 - procentBreak);
            if (num == "d")
                return Math.Round(Convert.ToDouble((BreakDinner.TotalMinutes) / Convert.ToDouble(PeriodWork.TotalMinutes))*100);
            if (num == "p")
                return Math.Round(Convert.ToDouble((BreakPause.TotalMinutes) / Convert.ToDouble(PeriodWork.TotalMinutes))*100);
            if (num == "m")
                return Math.Round(Convert.ToDouble((BreakMeeting.TotalMinutes) / Convert.ToDouble(PeriodWork.TotalMinutes)) * 100);
            if (num == "s")
                return Math.Round(Convert.ToDouble((BreakStudy.TotalMinutes) / Convert.ToDouble(PeriodWork.TotalMinutes)) * 100);
            if (num == "n")
                return Math.Round(Convert.ToDouble((BreakNote.TotalMinutes) / Convert.ToDouble(PeriodWork.TotalMinutes)) * 100);
            if (num == "dc")
                return Math.Round(Convert.ToDouble((BreakDoctor.TotalMinutes) / Convert.ToDouble(PeriodWork.TotalMinutes)) * 100);

            return 0.0;
        }

        public double WorkPeriods()
        {
            double workPeriods;
            double breakPeriods;
            workPeriods = Convert.ToDouble(PeriodWork.TotalMinutes);
            breakPeriods = Convert.ToDouble(BreakDinner.TotalMinutes) + Convert.ToDouble(BreakDoctor.TotalMinutes) + Convert.ToDouble(BreakMeeting.TotalMinutes) + Convert.ToDouble(BreakNote.TotalMinutes) + Convert.ToDouble(BreakPause.TotalMinutes) + Convert.ToDouble(BreakStudy.TotalMinutes);
            double procentWork = 0;
            double procentBreak = 0;

            procentWork = workPeriods;
            procentBreak = breakPeriods;

            procentBreak = Math.Round((procentBreak / procentWork) * 100);
            procentWork = Math.Round(100 - procentBreak);
            workPeriods = procentWork;
            //if (workPeriods > 0) { return workPeriods; } else { return 1.0; }
            return workPeriods;
        }
        public double BreakPeriods()
        {
            double workPeriods;
            double breakPeriods;
            workPeriods = Convert.ToDouble(PeriodWork.TotalMinutes);
            breakPeriods = Convert.ToDouble(BreakDinner.TotalMinutes) + Convert.ToDouble(BreakDoctor.TotalMinutes) + Convert.ToDouble(BreakMeeting.TotalMinutes) + Convert.ToDouble(BreakNote.TotalMinutes) + Convert.ToDouble(BreakPause.TotalMinutes) + Convert.ToDouble(BreakStudy.TotalMinutes);
            double procentWork = 0;
            double procentBreak = 0;

            procentWork = workPeriods;
            procentBreak = breakPeriods;

            procentBreak = Math.Round((procentBreak / procentWork) * 100);
            procentWork = Math.Round(100 - procentBreak);
            breakPeriods = procentBreak;
            //if (breakPeriods > 0) { return breakPeriods; } else { return 1.0; }
            return breakPeriods;
        }

        public TimeSpan PeriodBreaktxt()
        {
            
            TimeSpan breakPeriods;
            breakPeriods = BreakDinner + BreakDoctor + BreakMeeting + BreakNote + BreakPause + BreakStudy;
       
            return breakPeriods;
  
        }


    }
}
