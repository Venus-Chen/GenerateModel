using System.Collections.Generic;

namespace GenerateModel.Models
{
    /// <summary>
    /// Excel对象模型
    /// </summary>
    public class ExcelModel
    {
        /// <summary>
        /// Excel Sheet页A1列的值
        /// </summary>
        public string A1 { get; set; }

        /// <summary>
        /// Excel Sheet页B1列的值
        /// </summary>
        public string B1 { get; set; }

        /// <summary>
        /// Excel Sheet页C1列的值
        /// </summary>
        public string C1 { get; set; }

        /// <summary>
        /// Excel Sheet页D1列的值
        /// </summary>
        public string D1 { get; set; }

        /// <summary>
        /// Excel Sheet页E1列的值
        /// </summary>
        public string E1 { get; set; }

        /// <summary>
        /// Excel Sheet页F1列的值
        /// </summary>
        public string F1 { get; set; }

        /// <summary>
        /// 此Excel对象模型在Excel Sheet页中的索引
        /// </summary>
        public int Index { get; set; }

        /// <summary>
        /// Excel对象模型的父级
        /// </summary>
        public ExcelModel Parent { get; set; }

        /// <summary>
        /// Excel对象模型的子集
        /// </summary>
        public List<ExcelModel> Children { get; set; }
    }
}