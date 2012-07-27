#encoding: utf-8

require 'sequel'

UTF_2_IBM_CONVERTER = Encoding::Converter.new 'UTF-8', 'IBM866'
IBM_2_UTF_CONVERTER = Encoding::Converter.new 'IBM866', 'UTF-8'

# Конвертирует текст из Юникода в формат Microsoft Access.
def to_access(str)
  UTF_2_IBM_CONVERTER.convert str
end

# Конвертирует текст из формата Microsoft Access в Юникод.
def from_access(str)
  IBM_2_UTF_CONVERTER.convert str
end

RASSPASS = Sequel.ado(
  conn_string: "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=#{ARGV.shift}"
)

# puts RASSPASS[:Доставка].insert(:МетодДоставки => 'Привет') #.to_csv
# puts "!!!\n\n"

RASSPASS[:Доставка].each do |i|
  puts i.keys
  puts i[i.keys[1]]
#  i.each_value do |key|
#    puts from_access(key.to_s)
#    puts key.encoding
#    puts from_access(key.to_s) == 'МетодДоставки'
#  end
end
