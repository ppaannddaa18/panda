# 公共工具函数

def format_amount(amount, decimal=2):
    """格式化金额，添加千分位"""
    return f"{amount:,.{decimal}f}"


def chinese_amount(num):
    """数字转中文大写金额"""
    units = ['', '拾', '佰', '仟']
    large_units = ['元', '万', '亿']
    digits = ['零', '壹', '贰', '叁', '肆', '伍', '陆', '柒', '捌', '玖']

    if num == 0:
        return '零元整'

    # 简化实现，完整版可扩展
    result = []
    integer_part = int(num)
    decimal_part = round((num - integer_part) * 100)

    # 处理整数部分
    if integer_part > 0:
        str_int = str(integer_part)
        for i, digit in enumerate(str_int):
            pos = len(str_int) - i - 1
            if digit != '0':
                result.append(digits[int(digit)])
                result.append(units[pos % 4])
            elif result and result[-1] != '零':
                result.append('零')
        result.append('元')

    # 处理小数部分
    if decimal_part > 0:
        jiao = decimal_part // 10
        fen = decimal_part % 10
        if jiao > 0:
            result.append(digits[jiao] + '角')
        if fen > 0:
            result.append(digits[fen] + '分')
    else:
        result.append('整')

    return ''.join(result)