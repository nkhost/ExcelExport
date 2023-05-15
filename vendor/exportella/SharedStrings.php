<?php

namespace Exportella;

use Exception;

/**
 * Класс для работы с SharedStrings.xml
 */
class SharedStrings
{
  /**
   * Путь до sharedStrings.xml
   * @var string
   */
  protected $path;
  
  /**
   * Массив новых (добавленных) строк
   * @var array
   */
  protected $strings = [];
  
  /**
   * Индекс для вставки следующей строки
   * @var int
   */
  protected $index = 0;
  
  /**
   * Книга
   * @var Workbook
   */
  protected $workbook;
  
  /**
   * @param string $path Путь к файлу sharedStrings.xml
   * @param Workbook $workbook Книга
   */
  public function __construct(string $path, Workbook $workbook)
  {
    $this->path = $path;
    $this->workbook = $workbook;
  }
  
  /**
   * Загрузить sharedStrings.xml в память
   *
   * @return void
   * @throws Exception
   */
  public function load()
  {
    if (!file_exists($this->path)) {
      throw new \Exception('Не найден файл sharedStrings.xml. Шаблон должен содержать хотя бы одну ячейку с текстовым значением');
    }
    
    $matches = [];
    $fileContent = file_get_contents($this->path);
    preg_match_all('/<si[^>]*>(.+?)<\/si>/s', $fileContent, $matches);
    $this->strings = $matches[1];
    $this->index = count($this->strings);
  }
  
  /**
   * Сохранить sharedStrings.xml в файл
   *
   * @return void
   */
  public function save()
  {
    $file = fopen($this->path, 'wt');
    $stringsCount = count($this->strings);
    $header = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' .
      '<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="' . $stringsCount . '" uniqueCount="' . $stringsCount . '">';
    fputs($file, $header);
    
    if (!empty($this->strings)) {
      foreach ($this->strings as $string) {
        // Проверка наличия тегов (форматирования и стилей) внутри строки
        if (strpos($string, '<') !== false) {
          $sharedString = '<si>' . $string . '</si>';
        } else {
          // Если просто текст
          $sharedString = '<si><t>' . $string . '</t></si>';
        }
        fputs($file, $sharedString);
      }
      fputs($file, '</sst>');
    }
    
    fclose($file);
  }
  
  /**
   * Получить идентификатор строки
   *
   * @param string $string Строка
   * @param bool $ignoreDuplicates Не искать строку среди существующих, а записать как новую
   * @return int
   */
  public function getStringIndex(string $string, bool $ignoreDuplicates = true): int
  {
    if (!$ignoreDuplicates) {
      $index = array_search($string, $this->strings);
      if ($index) {
        return $index;
      }
    }
    
    $stringIndex = $this->index++;
    $this->strings[$stringIndex] = $string;
    
    return $stringIndex;
  }
  
  /**
   * Сформировать строку из подстрок. Каждая подстрока может иметь свой цвет и размер шрифта
   * @param string $string Массив строк
   * @param string|null $color Массив цветов в формате RGB (FF22AA) в соответствии с массивом строк. Null - не менять цвет
   * @param int|null $size Массив размеров в соответствии с массивом строк. Null - не менять размер
   * @return string Строка в XML формате
   */
  public static function customString(string $string, ?string $color = null, ?int $size=null): string
  {
    $result = '<r>';
    if (!empty($string) || !empty($color)) {
      $result .= '<rPr>';
      if ($color) {
        $result .= '<color rgb="' . $color . '"/>';
      }
      if ($size) {
        $result .= '<sz val="' . $size . '"/>';
      }
      $result .= '</rPr>';
    }
    $result .= '<t xml:space="preserve">' . $string . '</t></r>';
    return $result;
  }
}
