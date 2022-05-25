using System.Collections.Generic;

namespace GenerateModel.Helpers
{
    /// <summary>
    /// 变量级数定义
    /// </summary>
    public class NodeHelper
    {
        /// <summary>
        /// 一级变量
        /// </summary>
        public const string FirstNode = ">";

        /// <summary>
        /// 二级变量
        /// </summary>
        public const string SecondNode = ">>";

        /// <summary>
        /// 三级变量
        /// </summary>
        public const string ThirdNode = ">>>";

        /// <summary>
        /// 四级变量
        /// </summary>
        public const string FourthNode = ">>>>";

        /// <summary>
        /// 五级变量
        /// </summary>
        public const string FifthNode = ">>>>>";

        /// <summary>
        /// 六级变量
        /// </summary>
        public const string SixthNode = ">>>>>>";

        /// <summary>
        /// 七级变量
        /// </summary>
        public const string SeventhNode = ">>>>>>>";

        /// <summary>
        /// 八级变量
        /// </summary>
        public const string EighthNode = ">>>>>>>";

        /// <summary>
        /// 九级变量
        /// </summary>
        public const string NinthNode = ">>>>>>>>>";

        /// <summary>
        /// 十级变量
        /// </summary>
        public const string TenthNode = ">>>>>>>>>>";

        /// <summary>
        /// 变量层级集合
        /// </summary>
        public static readonly List<string> NodeList = new List<string>() 
        {
            FirstNode,
            SecondNode,
            ThirdNode,
            FourthNode,
            FifthNode,
            SixthNode,
            SeventhNode,
            EighthNode,
            NinthNode,
            TenthNode
        };
    }
}