<?php

namespace Exportella;

use \Exception;

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
   * Дескриптор открытого файла
   * @var resource
   */
  protected $file;
  
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
    $tag = '';          // Буфер для хранения последнего тега
    $inTag = false;     // Флаг - курсор чтения файла находится внутри тега
    $lastSiOffset = 0;  // Позиция последнего тега </si> а файле
    
    $this->file = fopen($this->path, 'r+t');
    if (!$this->file) {
      throw new Exception('Can\'t open sharedStrings.xml');
    }
    
    // Считываем xml-теги из файла
    while (!feof($this->file)) {
      $c = fgetc($this->file);
      
      if ($c === '<') { // Начало тега
        $inTag = true;
        $tag = '<';
      } elseif ($c === '>') { // Конец тега
        $inTag = false;
        $tag .= '>';
        
        if ($tag === '</si>') {
          $this->index++;
          $lastSiOffset = ftell($this->file); // Запоминаем положение последнего тега </si>
        }
      } elseif ($inTag) {
        $tag .= $c;
      }
    }
    
    fseek($this->file, $lastSiOffset); // Ставим курсор в файле после последнего тега </si>
  }
  
  /**
   * Сохранить sharedStrings.xml в файл
   *
   * @return void
   */
  public function save()
  {
    
    if (!$this->file) {
      return;
    }
    
    if (!empty($this->strings)) {
      
      foreach ($this->strings as $string) {
        fputs($this->file, '<si><t>' . $string . '</t></si>');
      }
      fputs($this->file, '</sst>');
    }
    
    fclose($this->file);
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
}
