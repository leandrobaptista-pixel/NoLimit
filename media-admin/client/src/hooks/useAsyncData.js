import { useCallback, useState } from 'react';

export function useAsyncData(initialValue) {
  const [data, setData] = useState(initialValue);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState('');

  const run = useCallback(async (task) => {
    setLoading(true);
    setError('');
    try {
      const result = await task();
      setData(result);
      return result;
    } catch (taskError) {
      setError(taskError.message || 'Unexpected error.');
      throw taskError;
    } finally {
      setLoading(false);
    }
  }, []);

  return {
    data,
    setData,
    loading,
    error,
    setError,
    run
  };
}
