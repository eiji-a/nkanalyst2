#!/usr/bin/ruby
#

require 'rubygems'
require 'yaml'
require 'spreadsheet'

MONTHS = ['07', '08', '09', '10', '11', '12', '01', '02', '03', '04', '05', '06']
ZOFFSET = [[2,  3], [5,  3], [8,  3], [11,  3], [14,  3], [17,  3],
           [2, 46], [5, 46], [8, 46], [11, 46], [14, 46], [17, 46]]
GOFFSET = [[2, 2], [2, 51], [2, 100]]
ZPARAMS = [['uriage', 0], ['siire', 3], ['sitenkan', 4], ['kisyu', 5],
           ['kimatu', 6], ['keihi', 37], ['zassyunyu', 38]]
GPARAMS = [['gzassyunyu', 40], ['gukerisoku', 41], ['siharisoku', 42],
           ['zason', 43], ['zeimodori', 44]] 
RUIKEI = 'ruikei'

def read_recipe
  if ARGV.size != 1
    STDERR.puts "Usage: get_zisseki.rb <recipe>"
    exit 1
  end

  $DIR = Dir.pwd
  open(ARGV[0], 'r') do |fp|
    $RECIPE = YAML::load(fp.read)
  end
end

def select_val(value)
  return 0.0 if value == nil
  return value if value.class == Float
  return 0.0 if value.value == nil
  return value.value
end

def get_z3val(sheet, offset, voff)
  val = Array.new
  val[0] = select_val(sheet[offset[1] + voff, offset[0] + 0])
  val[1] = select_val(sheet[offset[1] + voff, offset[0] + 1])
  val[2] = select_val(sheet[offset[1] + voff, offset[0] + 2])
  val
end

def get_g3val(sheet, hoff, offset)
  val = Array.new
  val[0] = select_val(sheet[GOFFSET[0][1] + offset, GOFFSET[0][0] + hoff]) 
  val[1] = select_val(sheet[GOFFSET[1][1] + offset, GOFFSET[1][0] + hoff]) 
  val[2] = select_val(sheet[GOFFSET[2][1] + offset, GOFFSET[2][0] + hoff]) 
  val
end

def read_month_zit(file, sheet, data)
  book = Spreadsheet.open file, 'rb'
  sheet = book.worksheet sheet
  MONTHS.each_index do |idx|
    mname = 'month' + MONTHS[idx]
    data[mname] = Hash.new
    ZPARAMS.each do |pa|
      data[mname][pa[0]] = get_z3val(sheet, ZOFFSET[idx], pa[1])
    end
  end 
  #book.close
end

def read_month_gai(file, sheet, data)
  book = Spreadsheet.open file, 'rb'
  sheet = book.worksheet sheet
  MONTHS.each_index do |idx|
    mname = 'month' + MONTHS[idx]
    GPARAMS.each do |pa|
      data[mname][pa[0]] = get_g3val(sheet, idx, pa[1])
    end
  end
  #book.close
end

def add_param(data, month, para)
  data[RUIKEI][para][0] += data[month][para][0]
  data[RUIKEI][para][1] += data[month][para][1]
  data[RUIKEI][para][2] += data[month][para][2]
end

def add_params(data, month)
  ZPARAMS.each do |zp|
    add_param(data, month, zp[0])
  end
  GPARAMS.each do |gp|
    add_param(data, month, gp[0])
  end
end

def set_zaiko_ruikei(data, mon)
  data[RUIKEI]['kisyu'][0] = data['month07']['kisyu'][0]
  data[RUIKEI]['kisyu'][1] = data['month07']['kisyu'][1]
  data[RUIKEI]['kisyu'][2] = data['month07']['kisyu'][2]
  data[RUIKEI]['kimatu'][0] = data['month' + mon]['kimatu'][0]
  data[RUIKEI]['kimatu'][1] = data['month' + mon]['kimatu'][1]
  data[RUIKEI]['kimatu'][2] = data['month' + mon]['kimatu'][2]
end

def calc_ruikei(data)
  data[RUIKEI] = Hash.new
  ZPARAMS.each do |zp|
    data[RUIKEI][zp[0]] = [0.0, 0.0, 0.0]
  end
  GPARAMS.each do |gp|
    data[RUIKEI][gp[0]] = [0.0, 0.0, 0.0]
  end  
  mon = sprintf("%02d", $RECIPE['month'])
  MONTHS.each do |m|
    add_params(data, 'month' + m)
    break if m == mon
  end
  set_zaiko_ruikei(data, mon)
end

def calc_zogen(data)
  data['zogen'] = Array.new
  data['zogen'][0] = data['kimatu'][0] - data['kisyu'][0]
  data['zogen'][1] = data['kimatu'][1] - data['kisyu'][1]
  data['zogen'][2] = data['kimatu'][2] - data['kisyu'][2]
end

def calc_arari(data)
  data['arari'] = Array.new
  data['arari'][0] = data['uriage'][0] - data['siire'][0] - data['sitenkan'][0] + data['zogen'][0]
  data['arari'][1] = data['uriage'][1] - data['siire'][1] - data['sitenkan'][1] + data['zogen'][1]
  data['arari'][2] = data['uriage'][2] - data['siire'][2] - data['sitenkan'][2] + data['zogen'][2]
end

def calc_eiri(data)
  data['eiri'] = Array.new
  data['eiri'][0] = data['arari'][0] - data['keihi'][0] + data['zassyunyu'][0]
  data['eiri'][1] = data['arari'][1] - data['keihi'][1] + data['zassyunyu'][1]
  data['eiri'][2] = data['arari'][2] - data['keihi'][2] + data['zassyunyu'][2]
end

def calc_gaisyukei(data)
  data['gaisyukei'] = Array.new
  data['gaisyukei'][0] = data['gzassyunyu'][0] + data['gukerisoku'][0]
  data['gaisyukei'][1] = data['gzassyunyu'][1] + data['gukerisoku'][1]
  data['gaisyukei'][2] = data['gzassyunyu'][2] + data['gukerisoku'][2]
end

def calc_gaisonkei(data)
  data['gaisonkei'] = Array.new
  data['gaisonkei'][0] = data['siharisoku'][0] + data['zason'][0]
  data['gaisonkei'][1] = data['siharisoku'][1] + data['zason'][1]
  data['gaisonkei'][2] = data['siharisoku'][2] + data['zason'][2]
end

def calc_sonekikei(data)
  data['sonekikei'] = Array.new
  data['sonekikei'][0] = data['gaisyukei'][0] - data['gaisonkei'][0]
  data['sonekikei'][1] = data['gaisyukei'][1] - data['gaisonkei'][1]
  data['sonekikei'][2] = data['gaisyukei'][2] - data['gaisonkei'][2]
end

def calc_keizyori(data)
  data['keizyori'] = Array.new
  data['keizyori'][0] = data['eiri'][0] + data['gaisyukei'][0] - data['gaisonkei'][0] + data['zeimodori'][0]
  data['keizyori'][1] = data['eiri'][1] + data['gaisyukei'][1] - data['gaisonkei'][1] + data['zeimodori'][1]
  data['keizyori'][2] = data['eiri'][2] + data['gaisyukei'][2] - data['gaisonkei'][2] + data['zeimodori'][2]
end

def calc_additionals(data)
  (MONTHS.map {|m| 'month' + m} << RUIKEI).each do |mon|
    dt = data[mon]
    calc_zogen(dt)
    calc_arari(dt)
    calc_eiri(dt)
    calc_gaisyukei(dt)
    calc_gaisonkei(dt)
    calc_sonekikei(dt)
    calc_keizyori(dt)
  end
end

def calc_eirikei(data)
  eirikei = [0.0, 0.0, 0.0]
  mon = sprintf("%02d", $RECIPE['month'])
  MONTHS.each do |m|
    month = 'month' + m
    eirikei[0] += data[month]['eiri'][0]
    eirikei[1] += data[month]['eiri'][1]
    eirikei[2] += data[month]['eiri'][2]
    data[month]['eirikei'] = Array.new
    data[month]['eirikei'][0] = eirikei[0]
    data[month]['eirikei'][1] = eirikei[1]
    data[month]['eirikei'][2] = eirikei[2]
  end
end

def set_cashflow(data)
  data[RUIKEI]['eirikei'] = Array.new
  data[RUIKEI]['eirikei'][0] = data[RUIKEI]['eiri'][0] - data[RUIKEI]['zogen'][0]
  data[RUIKEI]['eirikei'][1] = data[RUIKEI]['eiri'][1] - data[RUIKEI]['zogen'][1]
  data[RUIKEI]['eirikei'][2] = data[RUIKEI]['eiri'][2] - data[RUIKEI]['zogen'][2]

end

def gen_siten(name, para)
  puts "NAME:" + name
  puts para

  open(name + '.dat', 'w') do |fp|
    siten_dat = Hash.new
    siten_dat['year'] = $RECIPE['year']
    siten_dat['month'] = $RECIPE['month']
    read_month_zit(para['zitfile'], para['zitsheet'], siten_dat)
    read_month_gai(para['gaifile'], para['gaisheet'], siten_dat)
    calc_ruikei(siten_dat)
    calc_additionals(siten_dat)
    calc_eirikei(siten_dat)
    set_cashflow(siten_dat)
    fp.puts(siten_dat.to_yaml)
  end
end

def get_zisseki
  $RECIPE['siten'].each_pair do |st, val|
    next if val['zitfile'] == nil
    gen_siten(st, val)
  end
end

def main
  read_recipe
  get_zisseki
end

main

