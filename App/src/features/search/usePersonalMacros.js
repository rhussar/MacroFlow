import { useEffect, useRef, useState } from 'react';
import { normalizeMacros } from '../../lib/search-data';

export const PERSONAL_WORKBOOK_NAME = 'PERSONAL.XLSB';

const INITIAL_PERSONAL_MACROS_STATE = {
  status: 'idle',
  macros: [],
  workbookFound: false,
  error: null
};

export function usePersonalMacros(searchData, selectedWorkbookName = null) {
  const [personalState, setPersonalState] = useState(INITIAL_PERSONAL_MACROS_STATE);
  const requestSequence = useRef(0);

  useEffect(() => {
    if (searchData?.status !== 'ready') {
      requestSequence.current += 1;
      setPersonalState(INITIAL_PERSONAL_MACROS_STATE);
      return;
    }

    const workbookNameToCheck = String(
      selectedWorkbookName || searchData?.workbook?.name || ''
    )
      .trim()
      .toUpperCase();
    if (workbookNameToCheck === PERSONAL_WORKBOOK_NAME) {
      requestSequence.current += 1;
      setPersonalState({
        status: 'ready',
        macros: [],
        workbookFound: true,
        error: null
      });
      return;
    }

    const proceduresByWorkbookApi = window.excel?.vba?.proceduresByWorkbook;
    if (!proceduresByWorkbookApi) {
      setPersonalState({
        status: 'error',
        macros: [],
        workbookFound: false,
        error: { message: 'PERSONAL.XLSB macro API is unavailable.' }
      });
      return;
    }

    let cancelled = false;
    const requestId = ++requestSequence.current;

    setPersonalState((previous) => ({
      ...previous,
      status: 'loading',
      error: null
    }));

    (async () => {
      try {
        const result = await proceduresByWorkbookApi({ workbookName: PERSONAL_WORKBOOK_NAME });
        if (cancelled || requestId !== requestSequence.current) {
          return;
        }

        if (!result?.success) {
          const message = String(result?.message || result?.error || 'Unable to load PERSONAL.XLSB macros.');
          setPersonalState({
            status: 'error',
            macros: [],
            workbookFound: false,
            error: { message }
          });
          return;
        }

        const workbookFound = result?.workbookFound !== false;
        const macros = workbookFound ? normalizeMacros(result?.procedures) : [];
        setPersonalState({
          status: 'ready',
          macros,
          workbookFound,
          error: null
        });
      } catch (error) {
        if (cancelled || requestId !== requestSequence.current) {
          return;
        }
        const message = error?.message ? String(error.message) : 'Unable to load PERSONAL.XLSB macros.';
        setPersonalState({
          status: 'error',
          macros: [],
          workbookFound: false,
          error: { message }
        });
      }
    })();

    return () => {
      cancelled = true;
    };
  }, [
    searchData?.status,
    searchData?.workbook?.name,
    searchData?.workbook?.path,
    selectedWorkbookName
  ]);

  return personalState;
}
