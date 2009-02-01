#!ruby
require 'rexml/document'
require 'excel'
require 'kconv'

module Gansuke
  class Item
    attr_reader :name , :subs
    def initialize(name);
      @name , @subs = name , []
    end
    def add(subitem)
      @subs << subitem
    end
    def to_s
      buf = @name
      @subs.each{|sub|
        buf << "\n\t" + sub.to_s
      }
      buf
    end
  end
  class SubItem
    attr_reader :name , :charts
    def initialize(name)
      @name , @charts = name , []
    end
    def add(chart)
      @charts << chart
    end
    def to_s
      @name + " " + @charts.inject(String.new){|buf,chart|
        buf << chart.to_s
      }
    end
  end
  class Chart
    attr_reader :sdate , :edate
    def initialize(sdate , edate)
      @sdate , @edate = sdate , edate
    end
    def stime
      Time.local(sdate[0,4].to_i , sdate[4,2].to_i , sdate[6,2].to_i)
    end
    def etime
      Time.local(edate[0,4].to_i , edate[4,2].to_i , edate[6,2].to_i)
    end
    def to_s
      @sdate + " " + @edate
    end
  end
  class XMLParser
    def self.parse(file)
      items = []
      xml = REXML::Document.new open(file).read
      xml.each_element('//itemlist/item'){|eitem|
        item = Item.new(eitem.get_elements('text')[0].text.strip)
        items << item
        eitem.each_element('subitemlist/subitem'){|esubitem|
          text = esubitem.get_elements('text')[0].text
          text = text ? text.strip : ""
          item.add(SubItem.new(text))
        }
        eitem.each_element('chartlist/chart'){|echart|
          sdate = echart.get_elements('startdate')[0].text
          if sdate =~ /^(.*?)\/(.*?)\/(.*?) /
            sdate = $1 + $2.rjust(2,"0") + $3.rjust(2,"0")
          else
            sdate = ""
          end
          edate = echart.get_elements('enddate')[0].text
          if edate =~ /^(.*?)\/(.*?)\/(.*?) /
            edate = $1 + $2.rjust(2,"0") + $3.rjust(2,"0")
          else
            edate = ""
          end
          line  = echart.get_elements('line')[0].text.to_i
          item.subs[line].add(Chart.new(sdate , edate))
        }
      }
      items
    end
  end
end


def set_date_color(row , pos , stime , etime)
  time  = stime.clone
  while time < etime
    day = time.day
    if time.wday == 0 || time.wday == 6
      row.set_background_color(pos + 1 + (time - stime)/24/60/60 , "#FF80FF")
    end
    time = time + 24 * 60 * 60
  end

end

file = ARGV[0]
from = ARGV[1]
to   = ARGV[2]

chart_start = Time.local(from[0,4].to_i , from[4,2].to_i , from[6,2].to_i)
chart_end   = Time.local(to[0,4].to_i , to[4,2].to_i , to[6,2].to_i)

items = Gansuke::XMLParser.parse(file)

begin
  excel = Excel.new
  sheet = excel.get_sheet_at(1)
  row   = sheet.row 
  

  time  = chart_start.clone
  month = time.month
  cells = -1
  from  = 3
  while time < chart_end
    if month != time.month
      row.merge(from , from + cells)
      row.set_format(from , "@")
      row.set_value(from , time.year.to_s + "/" + month.to_s.rjust(2,"0"))
      from = from + cells + 1
      month = time.month
      cells = -1
    end
    time = time + 24 * 60 * 60
    cells += 1
  end
  if cells != 0
    row.merge(from , from + cells)
    row.set_format(from , "@")
    row.set_value(from , time.year.to_s + "/" + month.to_s.rjust(2,"0"))
  end
  
  row.next
  time  = chart_start.clone
  pos   = 3
  while time < chart_end
    day = time.day
    row.set_value(pos , time.day)
    time = time + 24 * 60 * 60
    pos += 1
  end
  set_date_color(row , 2 , chart_start , chart_end)
  sheet.set_width(2 , 2   , 20)
  sheet.set_width(3 , 255 , 2)
  row.next




  items.each {|item|
    set_date_color(row , 2 , chart_start , chart_end)
    row.set_value(1 , item.name.tosjis)
    row.set_border_top(1 , 3 + (chart_end - chart_start)/24/60/60)
    row.next
    item.subs.each{|sub|
      set_date_color(row , 2 , chart_start , chart_end)
      row.set_value(2,sub.name.tosjis)
      sub.charts.each{|chart|
        time = chart_start.clone
        while time < chart_end
          color = nil
          if time.wday == 0 || time.wday == 6
            color = "#FF80FF"
          end
          if time >= chart.stime && time <= chart.etime
            color = "#66B3FF"
          end
          row.set_background_color(3 + (time - chart_start)/24/60/60 , color) if color
          time = time + 24 * 60 * 60
        end
      }
      row.next
      set_date_color(row , 2 , chart_start , chart_end)
      row.set_height(2.25)
      row.next
    }
  }
  excel.save_current("schedule.xls")
ensure
  excel.close if excel
end
