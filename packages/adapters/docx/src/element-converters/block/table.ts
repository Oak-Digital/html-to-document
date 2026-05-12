import {
  BorderStyle,
  type FileChild,
  Paragraph,
  Table,
  TableCell,
  TableRow,
  TextRun,
  VerticalAlign,
} from 'docx';
import {
  type AttributeElement,
  cascadeStyles,
  computeInheritedStyles,
  type DocumentElement,
  type GridCell,
  type Styles,
  type TableElement,
} from 'html-to-document-core';
import type { ElementConverterDependencies, IBlockConverter } from '../types';

type DocumentElementType = TableElement;

const DOCX_DEFAULT_BORDER = {
  style: BorderStyle.SINGLE,
  size: 4,
  color: 'auto',
};

const DOCX_NONE_BORDER = {
  style: BorderStyle.NONE,
  size: 0,
  color: 'auto',
};

const ensureZeroSizeForNoneBorders = (
  mappedStyles: Record<string, unknown>
): Record<string, unknown> => {
  const normalizedStyles = { ...mappedStyles };

  for (const key of ['borders', 'border'] as const) {
    const borderGroup = normalizedStyles[key];
    if (
      !borderGroup ||
      typeof borderGroup !== 'object' ||
      Array.isArray(borderGroup)
    ) {
      continue;
    }

    const nextBorderGroup = { ...(borderGroup as Record<string, unknown>) };
    let changed = false;

    for (const [direction, borderValue] of Object.entries(nextBorderGroup)) {
      if (
        !borderValue ||
        typeof borderValue !== 'object' ||
        Array.isArray(borderValue)
      ) {
        continue;
      }

      const side = borderValue as Record<string, unknown>;
      if (side.style === BorderStyle.NONE && side.size !== 0) {
        nextBorderGroup[direction] = {
          ...side,
          size: 0,
        };
        changed = true;
      }
    }

    if (changed) {
      normalizedStyles[key] = nextBorderGroup;
    }
  }

  return normalizedStyles;
};

const hasNoneBorderStyle = (mappedStyles: Record<string, unknown>): boolean => {
  for (const key of ['borders', 'border'] as const) {
    const borderGroup = mappedStyles[key];
    if (
      !borderGroup ||
      typeof borderGroup !== 'object' ||
      Array.isArray(borderGroup)
    ) {
      continue;
    }

    for (const borderValue of Object.values(
      borderGroup as Record<string, unknown>
    )) {
      if (
        borderValue &&
        typeof borderValue === 'object' &&
        !Array.isArray(borderValue) &&
        (borderValue as Record<string, unknown>).style === BorderStyle.NONE
      ) {
        return true;
      }
    }
  }

  return false;
};

const ensureHiddenTableBorders = (
  mappedStyles: Record<string, unknown>
): Record<string, unknown> => {
  const normalizedStyles = ensureZeroSizeForNoneBorders(mappedStyles);

  return {
    ...normalizedStyles,
    borders: {
      top: { ...DOCX_NONE_BORDER },
      bottom: { ...DOCX_NONE_BORDER },
      left: { ...DOCX_NONE_BORDER },
      right: { ...DOCX_NONE_BORDER },
      insideHorizontal: { ...DOCX_NONE_BORDER },
      insideVertical: { ...DOCX_NONE_BORDER },
    },
  };
};

const getBorderGroup = (
  mappedStyles: Record<string, unknown>,
  key: 'border' | 'borders'
): Record<string, unknown> | undefined => {
  const value = mappedStyles[key];
  if (!value || typeof value !== 'object' || Array.isArray(value)) {
    return undefined;
  }

  return value as Record<string, unknown>;
};

const getBorderSide = (
  mappedStyles: Record<string, unknown>,
  direction: 'top' | 'right' | 'bottom' | 'left'
): Record<string, unknown> | undefined => {
  const borderGroup = getBorderGroup(mappedStyles, 'border');
  if (borderGroup?.[direction] && typeof borderGroup[direction] === 'object') {
    return borderGroup[direction] as Record<string, unknown>;
  }

  const bordersGroup = getBorderGroup(mappedStyles, 'borders');
  if (
    bordersGroup?.[direction] &&
    typeof bordersGroup[direction] === 'object'
  ) {
    return bordersGroup[direction] as Record<string, unknown>;
  }

  return undefined;
};

const isNoneBorderSide = (
  mappedStyles: Record<string, unknown>,
  direction: 'top' | 'right' | 'bottom' | 'left'
): boolean =>
  getBorderSide(mappedStyles, direction)?.style === BorderStyle.NONE;

const withExplicitCellBorders = (
  mappedStyles: Record<string, unknown>,
  border: Record<'top' | 'right' | 'bottom' | 'left', Record<string, unknown>>
): Record<string, unknown> => {
  const normalizedStyles = ensureZeroSizeForNoneBorders(mappedStyles);

  return {
    ...normalizedStyles,
    border,
    borders: border,
  };
};

export class TableConverter implements IBlockConverter<DocumentElementType> {
  public isMatch(element: DocumentElement): element is DocumentElementType {
    return element.type === 'table';
  }

  public async convertElement(
    dependencies: ElementConverterDependencies,
    element: TableElement,
    cascadedStyles?: Styles
  ): Promise<FileChild[]> {
    const { styleMapper, converter, defaultStyles, stylesheet, styleMeta } =
      dependencies;
    const captions: { side: string; paragraph: Paragraph }[] = [];

    // We filter the cascaded styles for the table scope
    const mergedStyles = {
      ...defaultStyles?.[element.type],
      ...stylesheet.getComputedStyles(element, cascadedStyles),
    };

    const cascadingStyles = cascadeStyles(
      mergedStyles,
      element.scope,
      styleMeta
    );

    // --- begin colgroup support ---
    let stylesCol: Record<string, unknown>[] = [];
    if (Array.isArray(element.metadata?.colgroup)) {
      const [colgroupMeta] = element.metadata.colgroup as AttributeElement[];
      stylesCol =
        (colgroupMeta?.metadata as { col: DocumentElement[] })?.col.map(
          (col) => {
            return ensureZeroSizeForNoneBorders(
              styleMapper.mapStyles(
                {
                  ...defaultStyles?.[col.type],
                  ...stylesheet.getComputedStyles(col, cascadingStyles),
                },
                col
              )
            );
          }
        ) ?? [];
    }

    if (Array.isArray(element.metadata?.caption)) {
      const caption = element.metadata.caption as AttributeElement[];
      captions.push(
        ...(await Promise.all(
          caption.map(async (c) => {
            const innerMergedStyles = {
              ...defaultStyles?.[c.type],
              ...stylesheet.getComputedStyles(c, cascadingStyles),
            };
            const innerCascadingStyles = cascadeStyles(
              innerMergedStyles,
              c.scope,
              styleMeta
            );
            return {
              side: (c.styles?.captionSide || 'top') as string,
              paragraph: new Paragraph({
                children: await converter.convertInline(
                  c,
                  stylesheet,
                  innerCascadingStyles
                ),
                ...styleMapper.mapStyles(innerMergedStyles, c),
              }),
            };
          })
        ))
      );
    }
    // --- end colgroup support ---
    const numRows = element.rows.length;

    let numCols = 0;

    for (const row of element.rows) {
      let colCount = 0;
      for (const cell of row.cells) {
        colCount += cell.colspan ? cell.colspan : 1;
      }
      numCols = Math.max(numCols, colCount);
    }

    const effectiveNumCols = Math.max(numCols, 1);

    const grid: (GridCell | null)[][] = Array.from({ length: numRows }, () => {
      return Array(effectiveNumCols).fill(null);
    });

    for (let i = 0; i < numRows; i++) {
      let colIndex = 0;
      const row = element.rows[i];
      if (!row) continue;

      for (const cell of row.cells) {
        const currentRow = grid[i]!;
        while (colIndex < effectiveNumCols && currentRow[colIndex] !== null)
          colIndex++;
        if (colIndex >= effectiveNumCols) break;
        const colSpan = cell.colspan || 1;
        const rowSpan = cell.rowspan || 1;

        currentRow[colIndex] = {
          cell,
          horizontal: false,
          verticalMerge: false,
          isMaster: true,
        };

        // Mark for horizontal merges
        for (let k = 1; k < colSpan; k++) {
          currentRow[colIndex + k] = {
            cell,
            horizontal: true,
            verticalMerge: false,
            isMaster: false,
          };
        }

        // Mark for vertical merge
        if (rowSpan > 1) {
          for (let r = i + 1; r < i + rowSpan && r < numRows; r++) {
            const nextRow = grid[r]!;
            nextRow[colIndex] = {
              cell,
              horizontal: false,
              verticalMerge: true,
              isMaster: false,
            };
          }
          colIndex += colSpan;
        }
      }
    }
    // Build the TableRows objects
    const tableRows: TableRow[] = [];
    const mappedTableStyles = ensureZeroSizeForNoneBorders(
      styleMapper.mapStyles({ ...mergedStyles, ...element.styles }, element)
    );
    const hiddenTableTop = isNoneBorderSide(mappedTableStyles, 'top');
    const hiddenTableRight = isNoneBorderSide(mappedTableStyles, 'right');
    const hiddenTableBottom = isNoneBorderSide(mappedTableStyles, 'bottom');
    const hiddenTableLeft = isNoneBorderSide(mappedTableStyles, 'left');
    const cellStyleCache = new Map<DocumentElement, Record<string, unknown>>();
    const getMappedCellStyles = (
      cell: DocumentElement
    ): Record<string, unknown> => {
      const cached = cellStyleCache.get(cell);
      if (cached) {
        return cached;
      }

      const mappedCellStyles = ensureZeroSizeForNoneBorders(
        styleMapper.mapStyles(
          {
            ...stylesheet.getMatchedStyles(cell),
            ...defaultStyles?.[cell.type],
            ...cell.styles,
          },
          cell
        )
      );

      cellStyleCache.set(cell, mappedCellStyles);
      return mappedCellStyles;
    };

    const hasAnyHiddenCellBorders = element.rows.some((row) => {
      return row.cells.some((cell) =>
        hasNoneBorderStyle(getMappedCellStyles(cell))
      );
    });

    const getNeighborMappedStyles = (
      rowIndex: number,
      columnIndex: number
    ): Record<string, unknown> | undefined => {
      const neighbor = grid[rowIndex]?.[columnIndex];
      if (!neighbor?.cell) {
        return undefined;
      }

      return getMappedCellStyles(neighbor.cell);
    };

    for (let i = 0; i < numRows; i++) {
      const cells: TableCell[] = [];
      let j = 0;
      while (j < effectiveNumCols) {
        const gridCell = grid[i]?.[j];
        if (!gridCell) {
          cells.push(
            new TableCell({
              verticalAlign: VerticalAlign.CENTER,
              ...(stylesCol[j] || {}),
              children: [
                new Paragraph({
                  children: [new TextRun({ text: '' })],
                }),
              ],
            })
          );
          j++;
        } else if (gridCell.horizontal) {
          j++;
        } else if (gridCell.verticalMerge) {
          cells.push(
            new TableCell({
              verticalMerge: 'continue',
              verticalAlign: VerticalAlign.CENTER,
              children: [
                new Paragraph({
                  children: [new TextRun({ text: '' })],
                }),
              ],
            })
          );
          j++;
        } else {
          const originalCell = gridCell.cell;
          const colSpan = originalCell?.colspan ? originalCell.colspan : 1;
          const rowSpan = originalCell?.rowspan ? originalCell.rowspan : 1;
          const verticalMerge = rowSpan > 1 ? 'restart' : undefined;
          const originalCellMatchedStyles = originalCell
            ? {
                ...defaultStyles?.[originalCell.type],
                ...stylesheet.getMatchedStyles(originalCell),
              }
            : {};
          const cellContent = originalCell
            ? converter.convertToBlocks({
                element: originalCell,
                stylesheet,
                cascadedStyles: computeInheritedStyles({
                  parentStyles: {
                    ...originalCellMatchedStyles,
                    ...stylesCol[j],
                    ...originalCell.styles,
                  },
                  parentScope: 'tableCell',
                  childScope: 'block',
                  metaRegistry: styleMeta,
                }),
                wrapInlineElements: (inlines) => {
                  return [
                    new Paragraph({
                      children: inlines,
                      ...styleMapper.mapStyles(
                        {
                          ...originalCellMatchedStyles,
                          ...stylesCol[j],
                          ...originalCell.styles,
                        },
                        originalCell
                      ),
                    }),
                  ];
                },
              })
            : [new Paragraph('')];

          let mappedCellStyles = originalCell
            ? getMappedCellStyles(originalCell)
            : {};

          if (hasAnyHiddenCellBorders && originalCell) {
            const topNeighborHidden = Array.from({ length: colSpan }).some(
              (_, offset) => {
                const neighborStyles = getNeighborMappedStyles(
                  i - 1,
                  j + offset
                );
                return neighborStyles
                  ? isNoneBorderSide(neighborStyles, 'bottom')
                  : false;
              }
            );
            const rightNeighborHidden = Array.from({ length: rowSpan }).some(
              (_, offset) => {
                const neighborStyles = getNeighborMappedStyles(
                  i + offset,
                  j + colSpan
                );
                return neighborStyles
                  ? isNoneBorderSide(neighborStyles, 'left')
                  : false;
              }
            );
            const bottomNeighborHidden = Array.from({ length: colSpan }).some(
              (_, offset) => {
                const neighborStyles = getNeighborMappedStyles(
                  i + rowSpan,
                  j + offset
                );
                return neighborStyles
                  ? isNoneBorderSide(neighborStyles, 'top')
                  : false;
              }
            );
            const leftNeighborHidden = Array.from({ length: rowSpan }).some(
              (_, offset) => {
                const neighborStyles = getNeighborMappedStyles(
                  i + offset,
                  j - 1
                );
                return neighborStyles
                  ? isNoneBorderSide(neighborStyles, 'right')
                  : false;
              }
            );

            mappedCellStyles = withExplicitCellBorders(mappedCellStyles, {
              top:
                isNoneBorderSide(mappedCellStyles, 'top') ||
                topNeighborHidden ||
                (i === 0 && hiddenTableTop)
                  ? { ...DOCX_NONE_BORDER }
                  : {
                      ...(getBorderSide(mappedCellStyles, 'top') ??
                        DOCX_DEFAULT_BORDER),
                    },
              right:
                isNoneBorderSide(mappedCellStyles, 'right') ||
                rightNeighborHidden ||
                (j + colSpan >= effectiveNumCols && hiddenTableRight)
                  ? { ...DOCX_NONE_BORDER }
                  : {
                      ...(getBorderSide(mappedCellStyles, 'right') ??
                        DOCX_DEFAULT_BORDER),
                    },
              bottom:
                isNoneBorderSide(mappedCellStyles, 'bottom') ||
                bottomNeighborHidden ||
                (i + rowSpan >= numRows && hiddenTableBottom)
                  ? { ...DOCX_NONE_BORDER }
                  : {
                      ...(getBorderSide(mappedCellStyles, 'bottom') ??
                        DOCX_DEFAULT_BORDER),
                    },
              left:
                isNoneBorderSide(mappedCellStyles, 'left') ||
                leftNeighborHidden ||
                (j === 0 && hiddenTableLeft)
                  ? { ...DOCX_NONE_BORDER }
                  : {
                      ...(getBorderSide(mappedCellStyles, 'left') ??
                        DOCX_DEFAULT_BORDER),
                    },
            });
          }

          cells.push(
            new TableCell({
              // TODO: make concurrent iterations
              children: await cellContent,
              columnSpan: colSpan > 1 ? colSpan : undefined,
              verticalMerge: verticalMerge,
              verticalAlign: VerticalAlign.CENTER,
              ...stylesCol[j],
              ...mappedCellStyles,
            })
          );
          j += colSpan;
        }
      }
      const rowElement = element.rows[i];
      // TODO: cascaded styles?
      const rowStyles = rowElement
        ? stylesheet.getComputedStyles(rowElement, undefined)
        : undefined;
      tableRows.push(
        new TableRow({
          children: cells,
          ...styleMapper.mapStyles(
            {
              ...(rowElement ? defaultStyles?.[rowElement?.type] : {}),
              ...rowStyles,
            },
            rowElement ?? element
          ),
        })
      );
    }

    // Drop table if it has no rows
    if (tableRows.length === 0) {
      return [];
    }

    const rawStyles = hasAnyHiddenCellBorders
      ? ensureHiddenTableBorders(mappedTableStyles)
      : mappedTableStyles;

    return [
      ...captions.filter((c) => c.side === 'top').map((c) => c.paragraph),
      new Table({
        ...rawStyles,
        rows: tableRows,
      }),
      ...captions.filter((c) => c.side === 'bottom').map((c) => c.paragraph),
    ];
  }
}
