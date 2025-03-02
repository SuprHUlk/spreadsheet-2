import React, { useState, ChangeEvent } from 'react';
import { FunctionTesterProps, MathFunctions, DataFunctions } from '../types';

const FunctionTester: React.FC<FunctionTesterProps> = ({ 
  onClose, 
  mathFunctions, 
  dataFunctions,
  onSelect 
}) => {
  const [selectedFunction, setSelectedFunction] = useState<string>('');
  const [testInput, setTestInput] = useState<string>('');
  const [result, setResult] = useState<any>(null);
  const [error, setError] = useState<string | null>(null);

  const allFunctions: { [key: string]: Function } = {
    ...mathFunctions,
    ...dataFunctions
  };

  const handleTest = (): void => {
    try {
      setError(null);
      let inputData: any[];
      
      // Parse input data
      try {
        // Try to parse as JSON first
        inputData = JSON.parse(`[${testInput}]`);
      } catch {
        // If JSON parse fails, split by commas
        inputData = testInput.split(',').map(val => {
          const num = parseFloat(val.trim());
          return isNaN(num) ? val.trim() : num;
        });
      }

      // Special handling for FIND_AND_REPLACE
      if (selectedFunction === 'FIND_AND_REPLACE') {
        const [text, find, replace] = inputData;
        const testResult = dataFunctions.FIND_AND_REPLACE({ 
          value: text, 
          find, 
          replace 
        });
        setResult(testResult);
        return;
      }

      // Execute the selected function
      const testResult = allFunctions[selectedFunction](inputData);
      setResult(testResult);

      // If test is successful, create the formula string
      if (testResult !== null && onSelect) {
        const formula = `${selectedFunction}(${testInput})`;
        onSelect(formula);
      }
    } catch (err) {
      setError(err instanceof Error ? err.message : String(err));
      setResult(null);
    }
  };

  return (
    <div className="function-tester">
      <div className="tester-header">
        <h3>Function Tester</h3>
        <button onClick={onClose} className="close-btn">×</button>
      </div>

      <div className="tester-content">
        <div className="input-group">
          <label>Select Function:</label>
          <select 
            value={selectedFunction} 
            onChange={(e: ChangeEvent<HTMLSelectElement>) => setSelectedFunction(e.target.value)}
          >
            <option value="">Choose a function...</option>
            <optgroup label="Mathematical Functions">
              <option value="SUM">SUM</option>
              <option value="AVERAGE">AVERAGE</option>
              <option value="COUNT">COUNT</option>
              <option value="MAX">MAX</option>
              <option value="MIN">MIN</option>
            </optgroup>
            <optgroup label="Data Functions">
              <option value="TRIM">TRIM</option>
              <option value="UPPER">UPPER</option>
              <option value="LOWER">LOWER</option>
              <option value="REMOVE_DUPLICATES">REMOVE_DUPLICATES</option>
              <option value="FIND_AND_REPLACE">FIND_AND_REPLACE</option>
            </optgroup>
          </select>
        </div>

        <div className="input-group">
          <label>Test Input:</label>
          <div className="input-help">
            {selectedFunction === 'FIND_AND_REPLACE' ? (
              "Enter: text, find, replace (comma-separated)"
            ) : (
              "Enter values (comma-separated or JSON array)"
            )}
          </div>
          <textarea
            value={testInput}
            onChange={(e: ChangeEvent<HTMLTextAreaElement>) => setTestInput(e.target.value)}
            placeholder={selectedFunction === 'FIND_AND_REPLACE' ? 
              'Example: Hello World, World, Earth' : 
              'Example: 1, 2, 3 or ["a", "b", "a"]'}
            rows={3}
          />
        </div>

        <button 
          onClick={handleTest}
          disabled={!selectedFunction || !testInput}
          className="test-btn"
        >
          Test Function
        </button>

        {(result !== null || error) && (
          <div className="result-section">
            <h4>Result:</h4>
            {error ? (
              <div className="error">{error}</div>
            ) : (
              <div className="result">
                {Array.isArray(result) ? 
                  JSON.stringify(result, null, 2) : 
                  String(result)
                }
              </div>
            )}
          </div>
        )}
      </div>
    </div>
  );
};

export default FunctionTester; 