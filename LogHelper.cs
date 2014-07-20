using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Diagnostics;

using log4net;
using log4net.Config;
using log4net.Appender;

namespace ConvertExcel
{
    /// <summary>
    /// 日志记录操作类（调用方式：LogHelper.Info("info")，为封装后Log4net类，五级日志，Debug->Info->WARN->Error->Fatal）
    /// </summary> 
    public class LogHelper
    {
        /// <summary>
        /// Debug委托
        /// </summary>
        /// <param name="message">日志信息</param>
        public delegate void DDebug(object message);

        /// <summary>
        /// Info委托
        /// </summary>
        /// <param name="message">日志信息</param>
        public delegate void DInfo(object message);

        /// <summary>
        /// Warn委托
        /// </summary>
        /// <param name="message">日志信息</param>
        public delegate void DWarn(object message);

        /// <summary>
        /// Error委托
        /// </summary>
        /// <param name="message">日志信息</param>
        public delegate void DError(object message);

        /// <summary>
        /// Fatal委托
        /// </summary>
        /// <param name="message">日志信息</param>
        public delegate void DFatal(object message);

        /// <summary>
        /// Debug
        /// </summary>
        public static DDebug Debug
        {
            get { return LogManager.GetLogger((new StackTrace()).GetFrame(1).GetMethod().DeclaringType).Debug; }
        }

        /// <summary>
        /// Info
        /// </summary>
        public static DInfo Info
        {
            get { return LogManager.GetLogger((new StackTrace()).GetFrame(1).GetMethod().DeclaringType).Info; }
        }

        /// <summary>
        /// Warn
        /// </summary>
        public static DWarn Warn
        {
            get { return LogManager.GetLogger((new StackTrace()).GetFrame(1).GetMethod().DeclaringType).Warn; }
        }

        /// <summary>
        /// Error
        /// </summary>
        public static DError Error
        {
            get { return LogManager.GetLogger((new StackTrace()).GetFrame(1).GetMethod().DeclaringType).Error; }
        }

        /// <summary>
        /// Fatal
        /// </summary>
        public static DFatal Fatal
        {
            get { return LogManager.GetLogger((new StackTrace()).GetFrame(1).GetMethod().DeclaringType).Fatal; }
        }

        /// <summary>
        /// 静态构造函数
        /// </summary>
        static LogHelper()
        {
            string path = string.Format("{0}log4net.config", AppDomain.CurrentDomain.BaseDirectory);
            if (File.Exists(path))
            {
                log4net.Config.XmlConfigurator.Configure(new FileInfo(path));
            }
            else
            {
                RollingFileAppender appender = new RollingFileAppender();
                appender.Name = "root";
                appender.File = "App_Data\\";
                appender.AppendToFile = true;
                appender.RollingStyle = RollingFileAppender.RollingMode.Composite;
                appender.DatePattern = "yyyyMMdd-HHmm\".txt\"";
                appender.MaximumFileSize = "2MB";
                appender.MaxSizeRollBackups = 10;
                log4net.Layout.PatternLayout layout = new log4net.Layout.PatternLayout("%d{yyyy-MM-dd HH:mm:ss,fff}[%t] %-5p [%c] : %m%n");
                appender.Layout = layout;
                BasicConfigurator.Configure(appender);
                appender.ActivateOptions();
            }
        }
    }
}
