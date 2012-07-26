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

mdb_filename = ARGV.shift

filename = 'Z:\\Sites\\Apps\\Nagruzka-2-RASSPASS\\test.mdb'

DB = Sequel.ado(
  :conn_string =>
      "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=#{filename}"
)

DB['Доставка'].select('Метод доставки').each do |i|
  i.each_key do |key|
    puts from_access(key.to_s)
    puts key.encoding
    puts from_access(key.to_s) == 'КодМетодаДоставки'
#    Encoding.name_list do |encoding_name|
#      puts encoding_name.class
#      converter = Encoding::Converter.new encoding_name, "UTF-8"
#      puts converter.convert key
#    end
  end
end
