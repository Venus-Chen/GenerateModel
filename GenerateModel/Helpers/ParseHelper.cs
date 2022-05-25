namespace GenerateModel.Helpers
{
    public class ParseHelper
    {
        public static bool ConvertToBool(string value)
        {
            var result = false;
            // 尝试去将string类型转换为bool类型
            if (bool.TryParse(value, out result))
            {
                // 转换成功，返回转换之后的值
                return result;
            }
            else
            {
                // 转换失败，返回默认值
                return result;
            }
        }
    }
}