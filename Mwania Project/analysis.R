## Microbiome diversity across soil, water, and skin samples
## Data: GlobalPatterns dataset (Caporaso et al. 2011), bundled with the
## phyloseq R package. Real 16S rRNA survey data across environments.

suppressMessages({
  library(phyloseq)
  library(DESeq2)
  library(ggplot2)
  library(dplyr)
})

data(GlobalPatterns)

# Subset down to the three environments we care about: soil, water, skin
ps <- subset_samples(GlobalPatterns,
                      SampleType %in% c("Soil", "Freshwater", "Ocean", "Skin"))
ps <- prune_taxa(taxa_sums(ps) > 0, ps)

# Collapse "Freshwater" and "Ocean" into one "Water" group for clarity
sample_data(ps)$env_group <- ifelse(sample_data(ps)$SampleType == "Soil", "Soil",
                              ifelse(sample_data(ps)$SampleType == "Skin", "Skin", "Water"))

# Keep only the most abundant taxa so the plot stays readable
top_taxa <- names(sort(taxa_sums(ps), decreasing = TRUE))[1:25]
ps_top   <- prune_taxa(top_taxa, ps)

# Convert to DESeq2 dataset using env_group as the design variable
dds <- phyloseq_to_deseq2(ps_top, ~ env_group)
dds <- estimateSizeFactors(dds, type = "poscounts")  # handles zero-inflated counts
dds <- DESeq(dds, fitType = "local")

res <- results(dds)
res_df <- as.data.frame(res) %>%
  mutate(taxon = rownames(res)) %>%
  filter(!is.na(padj))

# Pull in taxonomy (Phylum) and abundance-by-group for the dot plot
tax_df <- as.data.frame(tax_table(ps_top)) %>% mutate(taxon = rownames(.))
otu_df <- as.data.frame(otu_table(ps_top))
otu_df$taxon <- rownames(otu_df)

long_df <- otu_df %>%
  tidyr::pivot_longer(-taxon, names_to = "sample", values_to = "count") %>%
  left_join(
    data.frame(sample_data(ps_top), stringsAsFactors = FALSE) %>%
      mutate(sample = rownames(.)) %>%
      select(sample, env_group),
    by = "sample"
  ) %>%
  group_by(taxon, env_group) %>%
  summarise(mean_count = mean(count), .groups = "drop") %>%
  left_join(tax_df %>% select(taxon, Phylum), by = "taxon")

# Dot plot: mean abundance of top taxa across Soil / Water / Skin
p <- ggplot(long_df, aes(x = env_group, y = taxon, size = mean_count, color = Phylum)) +
  geom_point(alpha = 0.85) +
  scale_size(range = c(1, 10)) +
  theme_minimal(base_size = 11) +
  labs(title = "Top 25 Microbial Taxa Across Environments",
       subtitle = "GlobalPatterns dataset (Caporaso et al. 2011)",
       x = "Sample Type", y = "Taxon (OTU ID)", size = "Mean Abundance")

ggsave("microbiome_dotplot.png", p, width = 9, height = 8, dpi = 150)
cat("Saved plot to microbiome_dotplot.png\n")
cat("\nSignificant taxa (padj < 0.05):", sum(res_df$padj < 0.05, na.rm = TRUE), "of", nrow(res_df), "tested\n")

## ---------------------------------------------------------------
## Figure 2: Alpha diversity (Shannon index) across environments
## ---------------------------------------------------------------
alpha_df <- estimate_richness(ps, measures = "Shannon")
alpha_df$sample <- rownames(alpha_df)
alpha_df <- left_join(
  alpha_df,
  data.frame(sample_data(ps), stringsAsFactors = FALSE) %>%
    mutate(sample = rownames(.)) %>% select(sample, env_group),
  by = "sample"
)

p2 <- ggplot(alpha_df, aes(x = env_group, y = Shannon, fill = env_group)) +
  geom_boxplot(alpha = 0.7, outlier.shape = NA) +
  geom_jitter(width = 0.15, size = 2) +
  theme_minimal(base_size = 12) +
  theme(legend.position = "none") +
  labs(title = "Alpha Diversity (Shannon Index) by Environment",
       subtitle = "GlobalPatterns dataset (Caporaso et al. 2011)",
       x = "Sample Type", y = "Shannon Diversity Index")

ggsave("alpha_diversity_boxplot.png", p2, width = 7, height = 5, dpi = 150)
cat("Saved plot to alpha_diversity_boxplot.png\n")

# Kruskal-Wallis test: does Shannon diversity differ across environments?
kw <- kruskal.test(Shannon ~ env_group, data = alpha_df)
cat("\nKruskal-Wallis test on Shannon diversity across env_group:\n")
print(kw)

## ---------------------------------------------------------------
## Figure 3: Phylum-level relative abundance composition
## ---------------------------------------------------------------
ps_phylum <- tax_glom(ps, taxrank = "Phylum")
ps_phylum_rel <- transform_sample_counts(ps_phylum, function(x) x / sum(x))

phy_df <- psmelt(ps_phylum_rel) %>%
  group_by(env_group, Phylum) %>%
  summarise(mean_rel_abund = mean(Abundance), .groups = "drop") %>%
  group_by(env_group) %>%
  mutate(Phylum = ifelse(mean_rel_abund < 0.03, "Other (<3%)", as.character(Phylum))) %>%
  group_by(env_group, Phylum) %>%
  summarise(mean_rel_abund = sum(mean_rel_abund), .groups = "drop")

p3 <- ggplot(phy_df, aes(x = env_group, y = mean_rel_abund, fill = Phylum)) +
  geom_col(position = "stack") +
  theme_minimal(base_size = 12) +
  labs(title = "Mean Phylum-Level Composition by Environment",
       subtitle = "GlobalPatterns dataset (Caporaso et al. 2011)",
       x = "Sample Type", y = "Mean Relative Abundance") +
  scale_y_continuous(labels = scales::percent)

ggsave("phylum_composition_barplot.png", p3, width = 8, height = 6, dpi = 150)
cat("Saved plot to phylum_composition_barplot.png\n")

