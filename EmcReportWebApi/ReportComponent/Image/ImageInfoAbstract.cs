namespace EmcReportWebApi.ReportComponent.Image
{
    /// <summary>
    /// 图片
    /// </summary>
    public abstract class ImageInfoAbstract
    {
        /// <summary>
        /// 图片内容
        /// </summary>
        public string Content { get; set; }

        /// <summary>
        /// 图片名称
        /// </summary>
        public string ImageName { get; set; }

        /// <summary>
        /// 图片路径(路径+名称)
        /// </summary>
        public string ImageFileFullName { get; set; }
    }
}