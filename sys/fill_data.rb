#!/usr/bin/ruby
#

require 'yaml'
require 'rubygems'
require 'spreadsheet'

HOFFSET = [['uriage', 1], ['siire', 2], ['sitenkan', 3], ['kisyu', 4],
           ['kimatu', 5], ['zappi', 9], ['zassyunyu', 11],
           ['gzassyunyu', 16], ['gukerisoku', 17], ['siharisoku', 19],
           ['zason', 20], ['zeimodori', 23]]

MONTHS = ['07', '08', '09', '10', '11', '12', '01', '02', '03', '04', '05', '06']

def read_recipe
  if ARGV.size != 1
    STDERR.puts "Usage: fill_data.rb <recipe>"
    exit 1
  end

  $DIR = Dir.pwd
  open(ARGV[0], 'r') do |fp|
    $RECIPE = YAML::load(fp.read)
  end

  $SITEN = Hash.new
end

def read_data_file(key)
  begin
    open(key + '.dat') do |fp|
      $SITEN[key] = YAML::load(fp.read)
    end
  rescue StandardError => e
    STDERR.puts "READ ERROR (SITEN DATA): #{e.message}"
  end
end

def read_siten_data
  $RECIPE['siten'].each_pair do |key, val|
    puts "READ: #{key}"
    read_data_file(key)
  end
end

def read_offset(sheet)
  offsets = Hash.new
  $RECIPE['siten'].each_pair do |key, val|
    cols = sheet.column(0)
    idx = 0
    cols.each do |c|
      if val['name'] == c
        offsets[key] = idx + 3
        break
      end
      idx += 1
    end
  end
  offsets
end

def put_monthly(sheet, offset, mval)
  HOFFSET.each do |p|
    sheet[offset + 0, p[1]] = mval[p[0]][0]
    sheet[offset + 1, p[1]] = mval[p[0]][1]
    sheet[offset + 2, p[1]] = mval[p[0]][2]
  end
end

def put_values(sheet, offset, val)
  MONTHS.each_index do |i|
    #put_monthly(sheet, offset + 3 * i, val['month' + MONTHS[i]])
  end
end

def put_siten(sheet, offsets)
  $SITEN.each_pair do |s, val|
    put_values(sheet, offsets[s], val)
  end
end

def fill_data(sheet)
  offsets = read_offset sheet
  put_siten sheet, offsets
end

def fill_bunseki
  begin
    book = Spreadsheet.open "../" + $RECIPE['bunsekifile']
    sheet = book.worksheet sprintf("%02d", $RECIPE['month'])
    fill_data sheet

    sheet[4,2] = 12345
    book.write "../new-" + $RECIPE['bunsekifile']
  rescue StandardError => e
    STDERR.puts "FILE ERROR (#{$RECIPE['bunsekifile']}): #{e.message}"
    exit 1
  end

end

def main
  read_recipe
  read_siten_data
  fill_bunseki
end

main

