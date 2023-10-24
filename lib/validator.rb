#! /usr/bin/ruby

load 'powerpoint'

presentation = Powerpoint::ReadPresentation.new "~/Downloads/beyond-human_full_trend 3.pptx"
content_types_xml = Nokogiri::XML::Document.parse(presentation.files.file.open presentation.content_types_file)
presentation_rels_xml = Nokogiri::XML::Document.parse(presentation.files.file.open presentation.presentation_rels_file)

puts "Checking content type definitions exist"
content_types_xml.css('Override').each do |part|
  puts "#{part['PartName']} missing" unless presentation.files.file.open part['PartName'].gsub('/ppt','ppt')
end

puts "Checking presentation elements exist"
presentation_rels_xml.css('Relationship').each do |node|
  puts node['Target']
  puts "#{node['Target']}" unless presentation.files.file.open "ppt/#{node['Target']}"
end

puts "Checking slide rels"
presentation.slides.each do |slide|
  slide.raw_relation_content.css('Relationship').each do |node|
    puts "#{node['Target']}" unless presentation.files.file.open node['Target'].gsub('../','ppt/')
  end
end