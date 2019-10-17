格式使用说明
=======
    注：所有#注释行和空白行为忽略内容

* format.txt

    1. 列格式：
    
        对每个sheet表的各列进行格式设置， `[]`中表示该sheet表的名字, 对应后面为表头和对应列的格式，
        支持的列格式有:
        'width','hidden','level','collapsed','font_name','font_size','font_color','bold','italic','underline','font_strikeout',
        'font_script','num_format','locked','align','valign','rotation','text_wrap','reading_order','text_justlast','center_across',
        'indent','shrink','pattern','bg_color','fg_color','border','bottom','top','left','right','border_color','bottom_color',
        'top_color','left_color','right_color'
        表头和所有的值都无需引号, 值仅支持整数和字符串, 所有格式使用说明参见[python xlsxwriter format](https://xlsxwriter.readthedocs.io/format.html#format)

    2. 行格式：
    
        行格式可对每个sheet表的偶数列或奇数列进行颜色显示,标题行是否加粗，标题行是否自动过滤模式，如需设置，可按如下格式：
        - even_line_color(或single_line_color)= `colorValue`, colorValue无需引号，支持颜色英文或者RGB颜色码, 不设置则默认无颜色显示。
        - bold_header = `value`, value为0或False，则不加粗，若不设置或设置为1或True则使用加粗模式
        - auto_filter = `value`, value为0或False，则不自动过滤，若不设置或设置为1或True则自动添加过滤格式
                
    行格式用``[[row]]``标识，列格式用``[[column]]``标识，不指定可无需填写

* rule.txt

    对每个sheet表需要的列进行条件过滤，`[]`中表示该sheet表的名字, 对应后面为列的过滤规则，
    规则分为三步分"{COL}"{OP} {VAL}
    1. COL为列名，注意其被引号包起来。
    2. OP 为操作符，仅允许 ## 和 @@ 两种操作符，
        - 若为##则代表允许该列出现的值，后面VAL即为选取的值，VAL需为列表形式，使用()括起来，其中每个选取的值都需要添加引号
        - 若为@@代表该列的过滤规则，和面VAL即为过滤的规则，VAL格式为一个表达式或者两个表达式，最多只支持两个表达式。表达式可以用and、or、&&、||连接，表示与或关系。每个表达式中用x代表值，形式同python表达式语法，`所有符号或者值之间必须用空格分开`
        
* files.txt

    所输入的sheet表和对应的文件路径，sheet表的名字用`[]`括起来，后面跟文件路径，路径为相对`-b`参数指定的路径，下一行则指定文件内容解析的分割符, tsv表示"\t"分割，csv表示","分割，不指定则默认csv格式解析。

