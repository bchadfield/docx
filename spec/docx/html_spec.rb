require 'docx/html'
require 'test/unit'

class TestHtml < Test::Unit::TestCase
  def test_basic_conversion
    expected =
      '<!DOCTYPE html>'\
      '<html>'\
      '<head>'\
        '<title>basic</title>'\
      '</head>'\
      '<body>'\
        '<p>hello</p>'\
        '<p>world</p>'\
      '</body>'\
      '</html>'
    actual = Docx::Document.open('spec/fixtures/basic.docx').to_html(title: 'basic')
    assert_equal expected, actual
  end
  
  def test_conversion_with_formatting
    expected =
      '<!DOCTYPE html>'\
      '<html>'\
      '<head>'\
        '<title>formatting</title>'\
      '</head>'\
      '<body>'\
        '<p>Normal</p>'\
        '<p><em>Italic</em></p>'\
        '<p><strong>Bold</strong></p>'\
        '<p><span style="text-decoration: underline;">Underline</span></p>'\
        '<p>Normal</p>'\
        '<p>'\
          'This is a sentence with '\
          '<span style="text-decoration: underline;"><strong><em>all</em></strong></span> '\
          'formatting options in the middle of the sentence.'\
        '</p>'\
        '<h1>Heading 1</h1>'\
        '<h2>Heading 2</h2>'\
        '<h3>Heading 3</h3>'\
        '<h4>Heading 4</h4>'\
        '<h5>Heading 5</h5>'\
        '<h6>Heading 6</h6>'\
        '<p>A link</p>'\
        '</body>'\
        '</html>'
    actual = Docx::Document.open('spec/fixtures/formatting.docx').to_html(title: 'formatting')
    assert_equal expected, actual
  end
end
