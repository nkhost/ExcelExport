<?php

namespace ExcelExport;

use \Exception;

/**
 * Лист Xlsx документа - sheet_.xml
 */
class Worksheet
{
  /**
   * Путь до sheet_.xml
   * @var string
   */
  protected $path;
  
  /**
   * Содержимое sheet_.xml
   * @var string
   */
  protected $xml;
  
  /**
   * Объект Xlsx документа
   * @var Workbook
   */
  protected $workbook;
  
  /**
   * Объект для работы со строковыми значениями ячеек
   * @var SharedStrings
   */
  protected $sharedStrings;
  
  /**
   * Дескриптор файла для записи в режиме вставки
   * @var resource|null
   */
  protected $insertFile;
  
  /**
   * Положение начала открывающего тега <row> в который будет производиться вставка
   * @var int
   */
  protected $rowTagOffset;
  
  /**
   * Положение конца закрывающего тега <row> в который будет производиться вставка
   * @var int
   */
  protected $rowCloseTagEndOffset;
  
  /**
   * Номер строки куда будет производиться вставка
   * @var int
   */
  protected $rowIndex;
  
  /**
   * Количество вставленных строк
   * @var int
   */
  protected $insertedRows;
  
  /**
   * Номера столбцов для вставки
   * @var string[]
   */
  protected $insertColumns;
  
  /**
   * Флаг активации режима вставки
   * @var bool
   */
  protected $insertModeWasActivated = false;
  
  /**
   * @param string $path Путь к sheet_.xml
   * @param Workbook $workbook Книга
   * @throws Exception
   */
  public function __construct(string $path, Workbook $workbook)
  {
    $this->path = $path;
    $this->xml = file_get_contents($path);
    $this->workbook = $workbook;
    $this->sharedStrings = $this->workbook->getSharedStrings();
  }
  
  /**
   * Вставить значение в ячейку (ячейка не должна быть пустой в шаблоне).
   * Нельзя использовать после активации режима вставки.
   *
   * @param string $coordinates Координаты ячейки
   * @param string|int|float $value Значение для вставки
   * @param int|null $styleIndex Идентификатор стиля ячейки. Оставить стиль без изменения - null, убрать все стили - 0.
   * @return bool
   * @throws Exception
   */
  public function setCellValue(string $coordinates, $value, ?int $styleIndex = null): bool
  {
    if ($this->insertModeWasActivated) {
      throw new Exception('Insert mode was activated');
    }
    
    $cellOpenTagOffset = 0;     // Начало открывающего тега <c>
    $cellOpenTagEndOffset = 0;  // Конец открывающего тега <c>
    $cellCloseEndOffset = 0;    // Конец закрывающего тега </c>
    $tagOpen = '';              // Текст открывающего тега <c>
    
    $result = '';               // Результирующее содержимое xml файла
    
    // Находим ячейку по параметру
    $paramPosition = strpos($this->xml, "r=\"$coordinates\"");
    
    if (!$paramPosition) {
      return false;
    }
    
    // Ищем начало открывающего тега ячейки
    for ($i = $paramPosition; $i >= 0; $i--) {
      if ($this->xml[$i] == '<') {
        $cellOpenTagOffset = $i;
        break;
      }
    }
    
    // Ищем конец открывающего тега ячейки
    for ($i = $cellOpenTagOffset; $i < strlen($this->xml); $i++) {
      if ($this->xml[$i] === '>') {
        $tagOpen = substr($this->xml, $cellOpenTagOffset, $i - $cellOpenTagOffset + 1);
        $cellOpenTagEndOffset = $cellOpenTagOffset + strlen($tagOpen);
        break;
      }
    }
    
    // Если тег открывающий и закрывающий одновременно
    if (substr($tagOpen, -2, 2) === '/>') {
      $cellCloseEndOffset = $cellOpenTagEndOffset; // Конец открывающего тега и есть конец закрывающего тега
    } else {
      $cellCloseEndOffset = strpos($this->xml, '</c>', $cellOpenTagEndOffset) + 4;
    }
    
    // Записываем в результирующую строку всё от начала до ячейки
    $result .= substr($this->xml, 0, $cellOpenTagOffset);
    
    // Определяем тип данных
    $type = '';
    if (is_string($value)) {
      $type = 't="s"';
      $v = '<v>' . $this->sharedStrings->getStringIndex($value) . '</v>';
    } elseif (is_numeric($value)) {
      $v = '<v>' . str_replace(',', '.', $value) . '</v>';
    } elseif ($value === null) {
      $v = '';
    } elseif ($value instanceof Formula) {
      $v = "<f>$value</f>";
    } else {
      throw new Exception('Wrong data type');
    }
    
    // Определяем стиль
    $style = '';
    if ($styleIndex === null) { // Оставляем прежний стиль
      preg_match('/s=\"([0-9]+)\"/', $tagOpen, $matches);
      if (!empty($matches)) {
        $style = $matches[0];
      }
    } elseif ($styleIndex > 0) { // Убираем все стили
      $style = "s=\"$styleIndex\"";
    }
    
    // Если значение = null, то ячейку удаляем
    if ($value !== null) {
      // Записываем в результирующую строку ячейку
      $result .= "<c r=\"$coordinates\" $style $type>$v</c>";
    }
    
    // Записываем в результирующую строку от искомой ячейки до конца документа
    $result .= substr($this->xml, $cellCloseEndOffset, strlen($this->xml) - $cellCloseEndOffset);
    
    $this->xml = $result; // Перезаписываем содержимое xml-файла
    
    return true;
  }
  
  /**
   * Получить идентификатор стиля ячейки
   *
   * @param string $coordinates Координаты ячейки
   * @return int|null Идентификатор стиля
   */
  public function getStyleIndex(string $coordinates): ?int
  {
    $tagOpen = '';          // Текст открывающего тега <c>
    $cellOpenTagOffset = 0; // Начало открывающего тега <c>
    
    // Находим ячейку по параметру
    $paramPosition = strpos($this->xml, "r=\"$coordinates\"");
    
    if (!$paramPosition) {
      return null;
    }
    
    // Ищем начало открывающего тега
    for ($i = $paramPosition; $i >= 0; $i--) {
      if ($this->xml[$i] == '<') {
        $cellOpenTagOffset = $i;
        break;
      }
    }
    
    // Ищем конец открывающего тега
    for ($i = $cellOpenTagOffset; $i < strlen($this->xml); $i++) {
      if ($this->xml[$i] === '>') {
        $tagOpen = substr($this->xml, $cellOpenTagOffset, $i - $cellOpenTagOffset + 1);
        break;
      }
    }
    
    preg_match('/s=\"([0-9]+)\"/', $tagOpen, $matches);
    if (!empty($matches)) {
      return (int)$matches[1];
    }
    
    return null;
  }
  
  /**
   * Сохранить изменённые ячейки.
   * Нельзя использовать после активации режима вставки.
   *
   * @return void
   * @throws Exception
   */
  public function saveCells()
  {
    if ($this->insertModeWasActivated) {
      throw new Exception('Insert mode was activated');
    }
    
    file_put_contents($this->path, $this->xml);
  }
  
  /**
   * Инициировать режим вставки
   *
   * @param int $rowId Идентификатор строки для вставки (Строка должна быть не пустой в шаблоне)
   * @param array $insertColumns Названия столбцов для вставки. Например, ['A', 'B', 'C'] или ['A', 'B', 'D', 'J']
   * @throws Exception
   */
  public function initRowsInserting(int $rowId, array $insertColumns): void
  {
    $this->rowTagOffset = 0;
    $this->rowCloseTagEndOffset = 0;
    
    // Если файл уже открыт
    if ($this->insertFile) {
      throw new Exception('File already opened');
    }
    
    $this->insertFile = fopen($this->path, 'wt');
    if (!$this->insertFile) {
      throw new Exception('Can\'t open file');
    }
  
    preg_match("/<row[^>]+r=\"$rowId\"/", $this->xml, $matches, PREG_OFFSET_CAPTURE);
    if (!isset($matches[0][1])) {
      throw new Exception('Row does not exists in sheet');
    }
    $this->rowTagOffset = $matches[0][1];
    
    $this->rowCloseTagEndOffset = strpos($this->xml, '</row>', $this->rowTagOffset) + 6;
    
    // Записываем в файл sheet_.xml Всё до вставляемого блока
    for ($i = 0; $i < $this->rowTagOffset; $i++) {
      fwrite($this->insertFile, $this->xml[$i], 1);
    }
    
    $this->insertedRows = 0;
    $this->rowIndex = $rowId;
    $this->insertColumns = $insertColumns;
    $this->insertModeWasActivated = true;
  }
  
  /**
   * Вставка строки
   *
   * @param array $dataList Значение ячеек внутри строки
   * @param array|null $stylesList Идентификаторы стилей ячеек
   * @param float|null $height Высота строки
   * @param bool $ignoreDuplicates Не искать строковое значение среди существующих, а записать как новую
   * @param bool $preserveXmlTags
   * @return void
   * @throws Exception
   */
  public function insertRow(array $dataList, ?array $stylesList = [], ?float $height = null, bool $ignoreDuplicates = true, bool $preserveXmlTags = true): void
  {
    if(!$this->insertFile){
      throw new Exception('File not open');
    }
    
    $ht = $height ? ' ht="' . number_format($height, 2, '.', '') . '" customHeight="1"' : '';
    $row = '<row r="' . $this->rowIndex . '"' . $ht . '>';
    for ($i = 0; $i < count($dataList); $i++) {
      if (is_string($dataList[$i])) {
        $t = ' t="s"';
        $v = '<v>' . $this->sharedStrings->getStringIndex($dataList[$i], $ignoreDuplicates) . '</v>';
      } elseif (is_numeric($dataList[$i])) {
        $t = '';
        $v = '<v>' . str_replace(',', '.', $dataList[$i]) . '</v>';
      } elseif ($dataList[$i] instanceof Formula) {
        $t = '';
        $v = '<f>' . $dataList[$i] . '</f>';
      } else {
        continue;
      }
  
      $s = !empty($stylesList) && isset($stylesList[$i]) ? ' s="' . $stylesList[$i] . '"' : '';
      if (isset($this->insertColumns[$i])) {
        $row .= '<c r="' . $this->insertColumns[$i] . $this->rowIndex . '"' . $s . $t . '>' . $v . '</c>';
      }
    }
    $row .= '</row>';
    fputs($this->insertFile, $row);
  
    $this->rowIndex++;
    $this->insertedRows++;
  }
  
  /**
   * Завершить режим вставки
   *
   * @return void
   */
  public function finishRowsInserting(): void
  {
    // Участок документа, который должен располагаться после вставляемого блока
    $tail = substr($this->xml, $this->rowCloseTagEndOffset, strlen($this->xml) - $this->rowCloseTagEndOffset);
    
    // Увеличиваем номера у строк после вставленного блока на количество вставленных строк
    $incrementRow = function ($matches) {
      $oldIndex = (int)$matches[1];
      $newRowIndex = $oldIndex + $this->insertedRows - 1;
      return str_replace("r=\"$oldIndex\"", "r=\"$newRowIndex\"", $matches[0]);
    };
    $tail = preg_replace_callback('/<row[^>]+r=\"([0-9]+)\"/', $incrementRow, $tail);
    
    // Увеличиваем номера строк у ячеек после вставленного блока на количество вставленных строк
    $incrementCell = function ($matches) {
      $column = $matches[1];
      $oldIndex = (int)$matches[2];
      $newRowIndex = $oldIndex + $this->insertedRows - 1;
      return str_replace("r=\"$column$oldIndex\"", "r=\"$column$newRowIndex\"", $matches[0]);
    };
    $tail = preg_replace_callback('/<c[^>]+r=\"([A-Z]+)([0-9]+)\"/', $incrementCell, $tail);
    
    // Записываем в файл sheet_.xml Всё после вставляемого блока
    fputs($this->insertFile, $tail);
    fclose($this->insertFile);
    $this->insertFile = false;
  }
}
