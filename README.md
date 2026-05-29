# docs-ooxml

個人的にooxmlを直で触る必要があったので、docsファイルをooxmlで書いてみるプロジェクト

(zipなんも分からん)

## 使用ライブラリ

- react
- lexical
- fflate

## パフォーマンス

github actionsでベンチマークを取ってみた結果が以下の通り。
もしかしたら、コード修正で変わってるかも。最新はgithub actionsの結果を見てください。

| Mode       | Scenario | Paragraphs | List Items | Strategy      | Samples | Iterations/sample | Median avg ms/run |    Min-Max avg |     vs naive | DOCX size |
| ---------- | -------- | ---------: | ---------: | ------------- | ------: | ----------------: | ----------------: | -------------: | -----------: | --------: |
| single-run | small    |         20 |         10 | naive         |      21 |                 1 |            0.7385 |  0.4056-1.2781 |     baseline |   13.5 KB |
| single-run | small    |         20 |         10 | optimized     |      21 |                 1 |            0.5078 |  0.2775-0.8980 | 31.2% faster |   13.5 KB |
| single-run | small    |         20 |         10 | fflate-store  |      21 |                 1 |            0.1674 |  0.1164-0.2560 | 77.3% faster |   13.5 KB |
| single-run | small    |         20 |         10 | fflate-stream |      21 |                 1 |            0.1766 |  0.1374-0.5011 | 76.1% faster |   13.6 KB |
| single-run | medium   |        200 |         50 | naive         |      21 |                 1 |            2.1280 |  1.8255-6.0620 |     baseline |   69.3 KB |
| single-run | medium   |        200 |         50 | optimized     |      21 |                 1 |            1.9013 |  1.6676-4.8709 | 10.7% faster |   69.3 KB |
| single-run | medium   |        200 |         50 | fflate-store  |      21 |                 1 |            0.4711 |  0.3030-1.8124 | 77.9% faster |   69.3 KB |
| single-run | medium   |        200 |         50 | fflate-stream |      21 |                 1 |            0.4465 |  0.3101-0.7184 | 79.0% faster |   69.5 KB |
| single-run | large    |       1000 |        120 | naive         |      21 |                 1 |            8.0556 | 7.7136-12.9499 |     baseline |  299.2 KB |
| single-run | large    |       1000 |        120 | optimized     |      21 |                 1 |            7.7688 |  7.5036-9.8145 |  3.6% faster |  299.2 KB |
| single-run | large    |       1000 |        120 | fflate-store  |      21 |                 1 |            1.5207 |  1.2668-3.9649 | 81.1% faster |  299.2 KB |
| single-run | large    |       1000 |        120 | fflate-stream |      21 |                 1 |            1.4421 |  1.3087-1.7998 | 82.1% faster |  299.3 KB |
| batch      | small    |         20 |         10 | naive         |       7 |               200 |            0.3411 |  0.3272-0.3634 |     baseline |   13.5 KB |
| batch      | small    |         20 |         10 | optimized     |       7 |               200 |            0.2260 |  0.2204-0.2393 | 33.7% faster |   13.5 KB |
| batch      | small    |         20 |         10 | fflate-store  |       7 |               200 |            0.0653 |  0.0619-0.0749 | 80.8% faster |   13.5 KB |
| batch      | small    |         20 |         10 | fflate-stream |       7 |               200 |            0.0742 |  0.0669-0.0795 | 78.2% faster |   13.6 KB |
| batch      | medium   |        200 |         50 | naive         |       7 |                60 |            2.0153 |  1.8812-2.0678 |     baseline |   69.3 KB |
| batch      | medium   |        200 |         50 | optimized     |       7 |                60 |            1.7739 |  1.7154-1.8979 | 12.0% faster |   69.3 KB |
| batch      | medium   |        200 |         50 | fflate-store  |       7 |                60 |            0.4109 |  0.3475-0.4842 | 79.6% faster |   69.3 KB |
| batch      | medium   |        200 |         50 | fflate-stream |       7 |                60 |            0.3870 |  0.3544-0.4727 | 80.8% faster |   69.5 KB |
| batch      | large    |       1000 |        120 | naive         |       7 |                15 |            8.9313 |  8.4117-9.2599 |     baseline |  299.2 KB |
| batch      | large    |       1000 |        120 | optimized     |       7 |                15 |            8.2713 |  7.7606-8.5117 |  7.4% faster |  299.2 KB |
| batch      | large    |       1000 |        120 | fflate-store  |       7 |                15 |            1.9061 |  1.7763-2.0214 | 78.7% faster |  299.2 KB |
| batch      | large    |       1000 |        120 | fflate-stream |       7 |                15 |            1.8545 |  1.6931-2.1895 | 79.2% faster |  299.3 KB |
