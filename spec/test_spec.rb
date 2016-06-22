require 'powerpoint'

describe 'Powerpoint parsing a sample PPTX file' do
  before(:all) do


    @what_content = {"global"=>{"title"=>"Global", "dataType"=>"fieldset", "items"=>{"downloadableInput"=>{"title"=>"Downloadable => ", "dataType"=>"checkBox", "disabled"=>true, "items"=>["as PowerPoint", "as PDF"], "value"=>["as PowerPoint", "as PDF"]}}}, "headers"=>{"title"=>"Header Titles", "dataType"=>"fieldset", "items"=>{"headerInput1"=>{"title"=>"Header 1 Title:", "dataType"=>"text", "value"=>"5 Years ago"}, "headerInput2"=>{"title"=>"Header 2 Title:", "dataType"=>"text", "value"=>"Now"}, "headerInput3"=>{"title"=>"Header 3 Title:", "dataType"=>"text", "value"=>"5 Years time"}}},"rowsManagerInput"=>{"title"=>"Row Manager", "dataType"=>"manager_items", "disabled"=>false, "value"=>[{"title"=>"Row", "dataType"=>"fieldset", "items"=>{"textInput1"=>{"title"=>"Column 1", "dataType"=>"text", "value"=>"1"}, "textInput2"=>{"title"=>"Column 2", "dataType"=>"text", "value"=>"2"}, "textInput3"=>{"title"=>"Column 3", "dataType"=>"text", "value"=>"3"}}}, {"title"=>"Row", "dataType"=>"fieldset", "items"=>{"textInput1"=>{"title"=>"Column 1", "dataType"=>"text", "value"=>"4"}, "textInput2"=>{"title"=>"Column 2", "dataType"=>"text", "value"=>"5"}, "textInput3"=>{"title"=>"Column 3", "dataType"=>"text", "value"=>"6"}}}, {"title"=>"Row", "dataType"=>"fieldset", "items"=>{"textInput1"=>{"title"=>"Column 1", "dataType"=>"text", "value"=>"7"}, "textInput2"=>{"title"=>"Column 2", "dataType"=>"text", "value"=>"8"}, "textInput3"=>{"title"=>"Column 3", "dataType"=>"text", "value"=>"9"}}}]}}
    @deck = Powerpoint::Presentation.new
    @deck.add_intro 'Bicycle Of the Mind', 'created by Steve Jobs'
    @deck.add_ff_trend_intro_slide 'Cashless Society', 'Contactless credit/debit cards, NFC- and web-enabled phones and digital wallets continue to transform the future of payment methods  -  with major implications for the way we will shop and interact with brands in the future.', 'samples/images/image4.jpeg'
    @deck.add_ff_trend_what_next_slide 'What will happen next', @what_content
    @deck.add_textual_slide 'Why Mac?', ['Its cool!', 'Its light!']
    @deck.add_textual_slide 'Why Iphone?', ['Its fast!', 'Its cheap!']
    @deck.add_pictorial_slide 'JPG Logo', 'samples/images/sample_png.png'
    @deck.add_text_picture_slide('Text Pic Split', 'samples/images/sample_png.png', content = ['Here is a string', 'here is another'])
    @deck.add_pictorial_slide 'PNG Logo', 'samples/images/sample_png.png'
    @deck.add_picture_description_slide('Pic Desc', 'samples/images/sample_png.png', content = ['Here is a string', 'here is another'])
    @deck.add_picture_description_slide('JPG Logo', 'samples/images/sample_jpg.jpg', content = ['descriptions'])
    @deck.add_pictorial_slide 'GIF Logo', 'samples/images/sample_gif.gif', {x: 124200, y: 3356451, cx: 2895600, cy: 1013460}
    @deck.add_textual_slide 'Why Android?', ['Its great!', 'Its sweet!']
    @deck.save 'samples/pptx/sample.pptx' # Examine the PPTX file
  end

  it 'Create a PPTX file successfully.' do
    #@deck.should_not be_nil
  end
end
