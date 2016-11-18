require 'zip/filesystem'
require 'nokogiri'

module Powerpoint
  class Powerpoint::ReadSlide

    attr_reader :presentation,
                :slide_number,
                :slide_number,
                :slide_file_name

    def initialize presentation, slide_xml_path
      @presentation = presentation
      @slide_xml_path = slide_xml_path
      @slide_number = extract_slide_number_from_path slide_xml_path
      @slide_notes_xml_path = "ppt/notesSlides/notesSlide#{@slide_number}.xml"
      @slide_file_name = extract_slide_file_name_from_path slide_xml_path

      parse_slide
      parse_slide_notes
      parse_relation
    end

    def parse_slide
      slide_doc = @presentation.files.file.open @slide_xml_path
      @slide_xml = Nokogiri::XML::Document.parse slide_doc
      slide_doc.close if slide_doc
    end

    def parse_slide_notes
      slide_notes_doc = @presentation.files.file.open @slide_notes_xml_path rescue nil
      @slide_notes_xml = Nokogiri::XML::Document.parse(slide_notes_doc) if slide_notes_doc
      slide_notes_doc.close if slide_notes_doc
    end

    def parse_relation
      @relation_xml_path = "ppt/slides/_rels/#{@slide_file_name}.rels"
      if @presentation.files.file.exist? @relation_xml_path
        relation_doc = @presentation.files.file.open @relation_xml_path rescue nil
        @relation_xml = Nokogiri::XML::Document.parse relation_doc
        relation_doc.close if relation_doc
      end
    end

    def content
      content_elements @slide_xml
    end

    def raw_content
      @slide_xml
    end

    def raw_relation_content
      @relation_xml
    end

    def notes_content
      content_elements @slide_notes_xml
    end

    def title
      title_elements = title_elements(@slide_xml)
      title_elements.join(" ") if title_elements.length > 0
    end

    def layout
      xml.css('Relationship').select{ |node| node_is(node,'slideLayout') }.first
    end

    def master
      node = layout
      layout_rel_xml = Nokogiri::XML::Document.parse @presentation.files.file.open(
        node['Target'].gsub('..','ppt').gsub('slideLayouts','slideLayouts/_rels').gsub('xml','xml.rels'))
      layout_rel_xml.css('Relationship').select{ |node| node_is(node, 'slideMaster') }.first
    end

    def layout_file
      node = layout
      open_package_file node['Target']
    end

    def master_file
      node = master
      open_package_file node['Target']
    end


    def images
      image_elements(@relation_xml)
        .map.each do |node|
          open_package_file node['Target']
        end
    end

    def charts
      files = chart_elements(@relation_xml)
        .map.each do |node|
          open_package_file node['Target']
      end
      files  += chart_elements(@relation_xml)
        .map.each do |node|
          open_package_file(
            node['Target'].gsub('charts','charts/_rels').gsub('xml','xml.rels'))
      end
      files
    end

    def embeddings
      embeds = nil
      chart_elements(@relation_xml).each do |node|
        rel_file = @presentation.files.file.open(
          node['Target'].gsub('..','ppt').gsub('charts','charts/_rels').gsub('xml','xml.rels'))
        zip_entry = rel_file.rewind
        relation_doc = @presentation.files.file.open zip_entry.name
        embed_xml = Nokogiri::XML::Document.parse relation_doc
        relation_doc.close
        if embeds.nil?
          embeds = embedding_elements(embed_xml)
            .map.each do |node|
              open_package_file node['Target']
          end
        else
          embeds += embedding_elements(embed_xml)
            .map.each do |node|
              open_package_file node['Target']
          end
        end
        rel_file.close
      end
      embeds
    end

    def notes
      files = note_elements(@relation_xml)
        .map.each do |node|
          open_package_file node['Target']
      end
      files  += note_elements(@relation_xml)
        .map.each do |node|
          open_package_file(
            node['Target'].gsub('notesSlides','notesSlides/_rels').gsub('xml','xml.rels'))
      end
      files
    end

    def slide_num
      @slide_xml_path.match(/slide([0-9]*)\.xml$/)[1].to_i
    end

    private

    def extract_slide_number_from_path path
      path.gsub('ppt/slides/slide', '').gsub('.xml', '').to_i
    end

    def extract_slide_file_name_from_path path
      path.gsub('ppt/slides/', '')
    end

    def title_elements(xml)
      shape_elements(xml).select{ |shape| element_is_title(shape) }
    end

    def content_elements(xml)
      xml.xpath('//a:t').collect{ |node| node.text }
    end

    def image_elements(xml)
      xml.css('Relationship').select{ |node| node_is?(node, 'image') }
    end

    def chart_elements(xml)
       xml.css('Relationship').select{ |node| node_is?(node, 'chart') }
    end

    def embedding_elements(xml)
      xml.css('Relationship').select{ |node| node_is?(node, 'package') }
    end

    def note_elements(xml)
      xml.css('Relationship').select{ |node| node_is?(node, 'notesSlide') }
    end

    def shape_elements(xml)
      xml.xpath('//p:sp')
    end

    def element_is_title(shape)
      shape.xpath('.//p:nvSpPr/p:nvPr/p:ph').select{ |prop| prop['type'] == 'title' || prop['type'] == 'ctrTitle' }.length > 0
    end

    def node_is?(node, pattern)
      node['Type'].include? pattern
    end

    def open_package_file(file)
      @presentation.files.file.open file.gsub('..', 'ppt')
    end
  end
end
