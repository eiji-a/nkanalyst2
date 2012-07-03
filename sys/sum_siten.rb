#!/usr/bin/ruby
#

require 'yaml'

MONTHS = ['07', '08', '09', '10', '11', '12', '01', '02', '03', '04', '05', '06']
PARAMS = ['uriage', 'siire', 'sitenkan', 'kisyu', 'kimatu', 'keihi',
          'zassyunyu', 'gzassyunyu', 'gukerisoku', 'siharisoku',
          'zason', 'zeimodori', 'zogen', 'arari', 'eiri', 'eirikei',
          'gaisyukei', 'gaisonkei', 'sonekikei', 'keizyori']
RUIKEI = 'ruikei'

def read_recipe
  if ARGV.size != 1
    STDERR.puts "Usage: sum_siten.rb <recipe>"
    exit 1
  end

  $DIR = Dir.pwd
  open(ARGV[0], 'r') do |fp|
    $RECIPE = YAML::load(fp.read)
  end

  $SITEN = Hash.new
end

def read_data_file(key)
  open(key + '.dat') do |fp|
    $SITEN[key] = YAML::load(fp.read)
  end
end

def read_siten_data
  $RECIPE['siten'].each_pair do |key, val|
    next if val['zitfile'] == nil
    puts "READ: #{key}"
    read_data_file(key)
  end
end

def calc_sum(list, mname, para)
  sum = [0.0, 0.0, 0.0]
  list.each do |s|
    st = $SITEN[s][mname][para]
    sum[0] += st[0]
    sum[1] += st[1]
    sum[2] += st[2]
  end
  sum
end

def sum_siten(siten, list)
  MONTHS.each do |mon|
    mname = 'month' + mon
    siten[mname] = Hash.new
    PARAMS.each do |para|
      siten[mname][para] = calc_sum(list, mname, para)
    end
  end
  siten[RUIKEI] = Hash.new
  PARAMS.each do |para|
    siten[RUIKEI][para] = calc_sum(list, RUIKEI, para)
  end
end

def write_siten_data(name, siten)
  open(name + '.dat', 'w') do |fp|
    fp.puts(siten.to_yaml)
  end
end

def sumation
  $RECIPE['siten'].each_pair do |key, val|
    next if val['sitenlist'] == nil
    puts "GENERATE: #{key}"
    siten = Hash.new
    siten['year'] = $RECIPE['year']
    siten['month'] = $RECIPE['month']
    sum_siten(siten, val['sitenlist'])
    write_siten_data(key, siten) 
  end
end

def main
  read_recipe
  read_siten_data
  sumation
end

main

