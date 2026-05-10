function out = ipsc_minimal_g1(stageNum, numComb, outputFile)
%IPSC_MINIMAL_G1 Generate a first-generation HD-DE candidate set for iPSC media.
%
%   out = ipsc_minimal_g1(stageNum, numComb, outputFile)
%
%   This is a dry-run helper for the migration task. It does not modify or call
%   the original Excel/COM/liquid-handler workflow. It uses a curated factor
%   space extracted from docs/iPSC*.xlsx, generates G1 target vectors (X) and
%   trial vectors (U), and writes coded and decoded candidate tables.
%
%   stageNum   Differentiation stage to optimize, 1 through 5. Default: 5.
%   numComb    Number of X vectors and U vectors. Default: max(12, 3*numFact),
%              capped by available unique non-sparse combinations.
%   outputFile Excel output filename. Default: iPSC_stage##_G01_candidates.xlsx.
%
%   The output score/feedback loop is intentionally not implemented here,
%   because the biological objective metric and batch-analysis export format
%   still need to be defined for the new experiment.

if nargin < 1 || isempty(stageNum)
    stageNum = 5;
end

factors = ipscFactorSpace(stageNum);
numFact = numel(factors);

maxUnique = countNonSparseCombinations(factors);
maxPairComb = floor(maxUnique / 2);
if nargin < 2 || isempty(numComb)
    numComb = max(12, 3*numFact);
end
if numComb > maxPairComb
    warning('ipsc_minimal_g1:numCombCapped', ...
        'Requested numComb=%d, but only %d X/U pairs can be kept unique in this non-sparse factor space. Capping numComb.', ...
        numComb, maxPairComb);
    numComb = maxPairComb;
end
if numComb < 4
    error('numComb must be at least 4 because DE mutation requires r1/r2/r3 distinct from the current vector.');
end

if nargin < 3 || isempty(outputFile)
    outputFile = sprintf('iPSC_stage%02d_G01_candidates.xlsx', stageNum);
end

rng(0, 'twister');
mutCon = 1;
crossCon = 0.5;

X = generateTargetVectors(factors, numComb);
U = generateTrialVectors(X, factors, mutCon, crossCon);

writeCandidateWorkbook(outputFile, stageNum, factors, X, U);

out = struct();
out.stageNum = stageNum;
out.numFact = numFact;
out.numComb = numComb;
out.X = X;
out.U = U;
out.factors = factors;
out.outputFile = outputFile;

fprintf('Generated iPSC stage %d G1 dry-run candidates: %d X + %d U vectors.\n', ...
    stageNum, numComb, numComb);
fprintf('Output written to %s\n', outputFile);
end

function factors = ipscFactorSpace(stageNum)
% Factor spaces are stage-specific to keep the original HD-DE assumption that
% one vector represents one media condition at one experimental phase.
switch stageNum
    case 1
        factors = [ ...
            factor('Activin A', 'ng/mL', [0 1 7.5 10 25]), ...
            factor('BMP4', 'ng/mL', [0 5 10 30 40 50 80]), ...
            factor('CHIR99021', 'uM', [0 1 2 3]), ...
            factor('FGF2', 'ng/mL', [0 10 20]), ...
            factor('IWP2', 'uM', [0 2]), ...
            factor('SCF', 'ng/mL', [0 20]), ...
            factor('VEGF', 'ng/mL', [0 50]), ...
            factor('WNTC59', 'uM', [0 3]), ...
            factor('Y-27632', 'uM', [0 1 10])];
    case 2
        factors = [ ...
            factor('BMP4', 'ng/mL', [0 5]), ...
            factor('FGF2', 'ng/mL', [0 5 20 25 100]), ...
            factor('IL-3', 'ng/mL', [0 25]), ...
            factor('IL-34', 'ng/mL', [0 25]), ...
            factor('IL-6', 'ng/mL', [0 20]), ...
            factor('Insulin', 'ug/mL', [0 5]), ...
            factor('M-CSF', 'ng/mL', [0 25 100]), ...
            factor('SB431542', 'uM', [0 10]), ...
            factor('SCF', 'ng/mL', [0 100 200]), ...
            factor('VEGF', 'ng/mL', [0 15 50 80]), ...
            factor('Y-27632', 'uM', [0 10])];
    case 3
        factors = [ ...
            factor('DKK-1', 'ng/mL', [0 30]), ...
            factor('FGF2', 'ng/mL', [0 10 50]), ...
            factor('FLT3L', 'ng/mL', [0 50]), ...
            factor('IL-3', 'ng/mL', [0 10 25 30 50]), ...
            factor('IL-34', 'ng/mL', [0 25]), ...
            factor('IL-6', 'ng/mL', [0 10 50]), ...
            factor('Insulin', 'ug/mL', [0 5]), ...
            factor('M-CSF', 'ng/mL', [0 25 50 100]), ...
            factor('SCF', 'ng/mL', [0 10 50 100]), ...
            factor('TPO', 'ng/mL', [0 5 30 50]), ...
            factor('VEGF', 'ng/mL', [0 10 50])];
    case 4
        factors = [ ...
            factor('FLT3L', 'ng/mL', [0 50]), ...
            factor('GM-CSF', 'ng/mL', [0 25]), ...
            factor('IL-3', 'ng/mL', [0 25]), ...
            factor('IL-34', 'ng/mL', [0 10 100]), ...
            factor('Insulin', 'ug/mL', [0 5]), ...
            factor('M-CSF', 'ng/mL', [0 5 10 20 50 100])];
    case 5
        factors = [ ...
            factor('GM-CSF', 'ng/mL', [0 10]), ...
            factor('IL-34', 'ng/mL', [0 10 100]), ...
            factor('M-CSF', 'ng/mL', [0 10 20 25]), ...
            factor('TGF-beta', 'ng/mL', [0 50])];
    otherwise
        error('stageNum must be an integer from 1 through 5.');
end
end

function f = factor(name, unit, levels)
f = struct('name', name, 'unit', unit, 'levels', levels);
end

function X = generateTargetVectors(factors, numComb)
numFact = numel(factors);
poolSize = max(1000, numComb * 50);
C = [];

while size(C, 1) < numComb
    Q = quasiRandom(poolSize, numFact);
    M = zeros(poolSize, numFact);
    for j = 1:numFact
        nd = numel(factors(j).levels);
        M(:, j) = floor(Q(:, j) * nd);
        M(M(:, j) > nd - 1, j) = nd - 1;
    end

    M(sum(M == 0, 2) > numFact / 2, :) = [];
    C = [C; M]; %#ok<AGROW>
    C = uniqueRowsFirst(C);

    poolSize = poolSize * 2;
    if poolSize > 500000 && size(C, 1) < numComb
        error('Unable to generate enough unique target vectors for this factor space.');
    end
end

idx = randperm(size(C, 1), numComb);
X = transpose(C(idx, :));
end

function Q = quasiRandom(n, d)
if exist('sobolset', 'file') == 2
    p = sobolset(d);
    p = scramble(p, 'MatousekAffineOwen');
    Q = net(p, n);
else
    Q = rand(n, d);
end
end

function U = generateTrialVectors(X, factors, mutCon, crossCon)
[numFact, numComb] = size(X);
U = zeros(numFact, numComb);
tabu = transpose(X);

for i = 1:numComb
    pool = setdiff(1:numComb, i);
    rp = pool(randperm(numComb - 1, 3));
    V = X(:, rp(1)) + mutCon * (X(:, rp(2)) - X(:, rp(3)));
    jrand = randi(numFact);

    for j = 1:numFact
        if rand <= crossCon || j == jrand
            U(j, i) = V(j);
        else
            U(j, i) = X(j, i);
        end
        U(j, i) = round(U(j, i));
        U(j, i) = max(0, min(U(j, i), numel(factors(j).levels) - 1));
    end

    tries = 0;
    while ismember(transpose(U(:, i)), tabu, 'rows') && tries < 25
        tries = tries + 1;
        for j = 1:numFact
            U(j, i) = randi(numel(factors(j).levels)) - 1;
        end
        if sum(U(:, i) == 0) > numFact / 2
            U(1, i) = min(1, numel(factors(1).levels) - 1);
        end
    end
    tabu = [tabu; transpose(U(:, i))]; %#ok<AGROW>
end
end

function writeCandidateWorkbook(outputFile, stageNum, factors, X, U)
numFact = numel(factors);
numComb = size(X, 2);
vectors = [X U];
types = [repmat({'X'}, numComb, 1); repmat({'U'}, numComb, 1)];

summary = { ...
    'stage', stageNum; ...
    'numFact', numFact; ...
    'numComb_X', numComb; ...
    'numComb_U', numComb; ...
    'score_status', 'not evaluated'; ...
    'objective_placeholder', 'define experiment-specific readout and normalization'};
xlswrite(outputFile, summary, 'Summary', 'A1');

maxLevels = 0;
for j = 1:numFact
    maxLevels = max(maxLevels, numel(factors(j).levels));
end
factorSheet = cell(numFact + 1, 3 + maxLevels);
factorSheet(1, 1:3) = {'factor_id', 'factor', 'unit'};
for k = 1:maxLevels
    factorSheet{1, 3 + k} = sprintf('level_%d', k - 1);
end
for j = 1:numFact
    factorSheet{j + 1, 1} = j;
    factorSheet{j + 1, 2} = factors(j).name;
    factorSheet{j + 1, 3} = factors(j).unit;
    for k = 1:numel(factors(j).levels)
        factorSheet{j + 1, 3 + k} = factors(j).levels(k);
    end
end
xlswrite(outputFile, factorSheet, 'FactorSpace', 'A1');

coded = cell(2 * numComb + 1, 2 + numFact);
coded(1, 1:2) = {'vector_type', 'vector_id'};
for j = 1:numFact
    coded{1, j + 2} = factors(j).name;
end
for i = 1:2 * numComb
    coded{i + 1, 1} = types{i};
    coded{i + 1, 2} = i - (strcmp(types{i}, 'U') * numComb);
    for j = 1:numFact
        coded{i + 1, j + 2} = vectors(j, i);
    end
end
xlswrite(outputFile, coded, 'CodedVectors', 'A1');

decoded = cell(2 * numComb + 1, 2 + numFact);
decoded(1, 1:2) = {'vector_type', 'vector_id'};
for j = 1:numFact
    decoded{1, j + 2} = factors(j).name;
end
for i = 1:2 * numComb
    decoded{i + 1, 1} = types{i};
    decoded{i + 1, 2} = i - (strcmp(types{i}, 'U') * numComb);
    for j = 1:numFact
        levelIndex = vectors(j, i) + 1;
        decoded{i + 1, j + 2} = sprintf('%g %s', ...
            factors(j).levels(levelIndex), factors(j).unit);
    end
end
xlswrite(outputFile, decoded, 'DecodedVectors', 'A1');

feedback = cell(2 * numComb + 3, 5);
feedback(1, :) = {'vector_type', 'vector_id', 'raw_readout', 'normalized_score', 'notes'};
feedback(2, :) = {'PC', 1, [], 1, 'positive control'};
feedback(3, :) = {'NC', 1, [], [], 'negative control'};
for i = 1:2 * numComb
    feedback{i + 3, 1} = types{i};
    feedback{i + 3, 2} = i - (strcmp(types{i}, 'U') * numComb);
end
xlswrite(outputFile, feedback, 'FeedbackTemplate', 'A1');
end

function C = uniqueRowsFirst(C)
[~, ia] = unique(C, 'rows', 'first');
ia = sort(ia);
C = C(ia, :);
end

function n = countNonSparseCombinations(factors)
numFact = numel(factors);
levels = zeros(1, numFact);
for j = 1:numFact
    levels(j) = numel(factors(j).levels);
end

total = prod(levels);
if total > 1000000
    n = total;
    return;
end

C = 0;
grid = cell(1, numFact);
for j = 1:numFact
    grid{j} = 0:levels(j) - 1;
end
idx = cell(1, numFact);
[idx{:}] = ndgrid(grid{:});
M = zeros(total, numFact);
for j = 1:numFact
    M(:, j) = idx{j}(:);
end
C = sum(sum(M == 0, 2) <= numFact / 2);
n = C;
end
