#!ruby
require 'win32ole'

module Sheet
  BORDER_LEFT   = 1
  BORDER_RIGHT  = 2
  BORDER_TOP    = 3
  BORDER_BOTTOM = 4

  def set_value(x , y , value)
    x_t = convert_numeric_to_s(x)
    # �z��œn���Ă����ꍇ
    if value.kind_of?(Array)
      for i in 0...value.length
        self.Cells.Item(y,x_t).Value = value[i]
        x_t.next!
      end
    else
      self.Cells.Item(y,x_t).Value = value
    end
  end
  #
  # �Z���̒l���擾����
  #
  def get_value(x , y)
    self.Cells.Item(y,x).Value
  end
  #
  # �Z���̃t�H�[�}�b�g���w�肷��
  # @ : ������
  #
  def set_format(x , y , format)
    self.Cells.Item(y,x).NumberFormatLocal = format
  end
  #
  # �񕝂�ݒ肷��
  # ��)
  # excel.set_width("A1" , "B1" , 100)
  #
  def set_width(from , to , width)
    from = convert_numeric_to_s(from) + "1"
    to   = convert_numeric_to_s(to)   + "1"

    self.Range(from , to).ColumnWidth = width
  end
  #
  #
  #
  def set_height(height , rowIndex)
    pos = "A" + rowIndex.to_s
    self.Range(pos).RowHeight = height
  end
  #
  #
  #
  def set_border_top(rowIndex , from , to)
    term = convert_numeric_to_s(from) + rowIndex.to_s 
    term << ":"
    term << convert_numeric_to_s(to) + rowIndex.to_s
    self.Range(term).Borders(BORDER_TOP).LineStyle = true
  end
  #
  #
  #
  def merge_cell(rowIndex , from , to)
    term = convert_numeric_to_s(from) + rowIndex.to_s 
    term << ":"
    term << convert_numeric_to_s(to) + rowIndex.to_s
    self.Range(term).MergeCells = true
  end
  #
  # �Z���̃t�H���g�F��ݒ肷��
  # color : 16�i�̐F��� 
  #         �� #FF0000
  #         �� #00FF00
  #         �� #0000FF
  #
  def set_color(x , y , color)
    x = convert_numeric_to_s(x)
    # �F���Z�b�g����
    self.Cells.Item(y,x).Font.Color = convert_color_tohex(color)
  end
  # 
  # �w�i�F��ݒ肷��
  # color : 16�i�̐F��� 
  #         �� #FF0000
  #         �� #00FF00
  #         �� #0000FF
  #
  def set_background_color(x , y , color)
    x = convert_numeric_to_s(x)
    # �w�i�F���Z�b�g����
    self.Cells.Item(y,x).Interior.Color = convert_color_tohex(color)
  end
  #
  # �s���
  #
  def row(index=1)
    return Row.new(self , index)
  end


  private
  def convert_color_tohex(color)
    # #�����
    if color =~ /^#/
      color = color.slice(1 , color.length)
    end
    # 6�o�C�g�����ꍇ��0
    if color.length != 6
      return 0
    end
    # R �� B �����ւ�
    color = color[4,2] + color[2,2] + color[0,2]
    # 16�i�ɖ߂��ĕԋp
    return color.hex
  end

  def convert_numeric_to_s(x)
    if x.kind_of?(Numeric) && x > 256
      throw Exception.new("256�ȏ�̐��l�͎w��ł��܂���")
    end
    # ���l�ϊ�
    if x.kind_of? Numeric
      tmp = "A"
      for i in 1...x
        tmp.next!
      end
      x = tmp
    end

    return x
  end
end

class Row

  def initialize(sheet , index=1)
    if index < 1
      throw Exception.new("�J�n�C���f�b�N�X�� 1 �ȏ�ł��B -> " + index)
    end
    @sheet = sheet
    @index = index
  end

  def next
    @index += 1
  end

  def has_nex?
    v = get_value(0,@index + 1)
    reurn !v || v != 22
  end

  def row_index
    @index
  end

  def set_index(index)
    @index = index
  end

  def set_border_top(from , to)
    @sheet.set_border_top(@index , from , to)
  end

  def merge(from , to)
    @sheet.merge_cell(@index , from , to)
  end

  def method_missing(method , *arg)
    #if method.to_s !~ /^set_/
    #  puts "method missing : [#{method}]"
    #  return
    #end
    arg.insert(1,@index)
    #@sheet.send(method , arg[0] , @index , arg[1])
    @sheet.send(method , *arg)
  end
end


class Excel
  def initialize(path=nil)
    begin
      # Excel�A�v���P�[�V�������쐬����
      @excel = WIN32OLE.new('Excel.Application')
      # ���[�N�u�b�N��ǉ�����
      if path
        @book  = @excel.Workbooks.open(path)
      else
        @book  = @excel.Workbooks.Add
      end
    rescue => e
      puts e
      close
    end
  end

  def add_sheet(name)
    # ���[�N�u�b�N�ɃV�[�g����ǉ�����
    @book.Worksheets.Add
    # �V�[�g�̖��O��ύX����
    @book.Worksheets(1).Name = name

    return @book.Worksheets(1).extend Sheet
  end

  def get_sheet(name)
    return @book.Worksheets.Item(name).extend Sheet
  end

  def get_sheet_at(index)
    return @book.Worksheets(index).extend Sheet
  end
  # �t�@�C����ۑ�����
  # path : �t�@�C����ۑ������΃p�X���w�肷��
  def save(path)
    @book.SaveAs(path)
  end
  def save_current(name)
    path = Dir.pwd + "/" + name
    if path =~ /^\/cygdrive\/(.*?)\//
      path = $1 + ":/" + path.sub(/^\/cygdrive\/(.*?)\//,"")
    end
    save(path)
  end

  def close
    # close workbook
    @book.Close if @book
    # quit excel application
    @excel.Quit if @excel
  end
end
