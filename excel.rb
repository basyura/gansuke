#!ruby
require 'win32ole'

module Sheet
  BORDER_LEFT   = 1
  BORDER_RIGHT  = 2
  BORDER_TOP    = 3
  BORDER_BOTTOM = 4

  def set_value(x , y , value)
    x_t = convert_numeric_to_s(x)
    # 配列で渡ってきた場合
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
  # セルの値を取得する
  #
  def get_value(x , y)
    self.Cells.Item(y,x).Value
  end
  #
  # セルのフォーマットを指定する
  # @ : 文字列
  #
  def set_format(x , y , format)
    self.Cells.Item(y,x).NumberFormatLocal = format
  end
  #
  # 列幅を設定する
  # 例)
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
  # セルのフォント色を設定する
  # color : 16進の色情報 
  #         赤 #FF0000
  #         緑 #00FF00
  #         青 #0000FF
  #
  def set_color(x , y , color)
    x = convert_numeric_to_s(x)
    # 色をセットする
    self.Cells.Item(y,x).Font.Color = convert_color_tohex(color)
  end
  # 
  # 背景色を設定する
  # color : 16進の色情報 
  #         赤 #FF0000
  #         緑 #00FF00
  #         青 #0000FF
  #
  def set_background_color(x , y , color)
    x = convert_numeric_to_s(x)
    # 背景色をセットする
    self.Cells.Item(y,x).Interior.Color = convert_color_tohex(color)
  end
  #
  # 行情報
  #
  def row(index=1)
    return Row.new(self , index)
  end


  private
  def convert_color_tohex(color)
    # #を削る
    if color =~ /^#/
      color = color.slice(1 , color.length)
    end
    # 6バイト無い場合は0
    if color.length != 6
      return 0
    end
    # R と B を入れ替え
    color = color[4,2] + color[2,2] + color[0,2]
    # 16進に戻して返却
    return color.hex
  end

  def convert_numeric_to_s(x)
    if x.kind_of?(Numeric) && x > 256
      throw Exception.new("256以上の数値は指定できません")
    end
    # 数値変換
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
      throw Exception.new("開始インデックスは 1 以上です。 -> " + index)
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
      # Excelアプリケーションを作成する
      @excel = WIN32OLE.new('Excel.Application')
      # ワークブックを追加する
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
    # ワークブックにシートをを追加する
    @book.Worksheets.Add
    # シートの名前を変更する
    @book.Worksheets(1).Name = name

    return @book.Worksheets(1).extend Sheet
  end

  def get_sheet(name)
    return @book.Worksheets.Item(name).extend Sheet
  end

  def get_sheet_at(index)
    return @book.Worksheets(index).extend Sheet
  end
  # ファイルを保存する
  # path : ファイルを保存する絶対パスを指定する
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
